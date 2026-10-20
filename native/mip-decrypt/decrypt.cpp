// mip-decrypt: unprotect one Purview-protected file via the MIP File SDK.
//
// This is the native companion to the ms-graph-mcp Node launcher
// (src/util/mip-decrypt.ts), which owns the protocol documented here:
//
//   argv:    mip-decrypt <input> <output>
//   exit 0: <output> holds the plaintext (a byte copy when the input was not
//           actually protected, so the caller can read uniformly)
//   exit 2: usage error (wrong argc, missing configuration)
//   exit 1: any decrypt failure
//   stderr: `[mip] ...` diagnostics; a `[mip] NO_RIGHTS: ...` line marks a
//           rights refusal so the caller can tell it from corruption.
//   stdout: unused. Token material is never printed anywhere.
//
// Configuration arrives by environment, never argv, so secrets stay out of the
// process table:
//
//   MIP_DECRYPT_CLIENT_ID   Entra application (client) ID for the MIP engine.
//   MIP_DECRYPT_USERNAME    UPN the engine acts as (engine identity).
//   MIP_DECRYPT_AUTHORITY   Authority URL. Defaults to the common endpoint.
//   MIP_DECRYPT_AUTH_HELPER Path to the token helper (see README: contract is
//                           `<helper> <client-id> <authority> <scope>` printing
//                           the access token on stdout, diagnostics on stderr).
//   MIP_DECRYPT_STORAGE     MIP storage directory. Defaults to a per-user temp
//                           subdir; the SDK needs a writable one.
//
// Build: ./build.sh (needs the MIP SDK; see README). No test in this repo
// executes this binary — the Node suite drives a fake sidecar instead — so
// keep the argv/exit/stderr contract above in sync with the launcher by hand.
#include <sys/wait.h>

#include <cstdio>
#include <cstdlib>
#include <filesystem>
#include <future>
#include <iostream>
#include <memory>
#include <string>

#include "mip/file/file_engine.h"
#include "mip/file/file_handler.h"
#include "mip/file/file_profile.h"
#include "mip/mip_context.h"

namespace {

constexpr char kDefaultAuthority[] = "https://login.microsoftonline.com/common";
constexpr char kAadrmResource[] = "https://aadrm.com";
constexpr char kAadrmScope[] = "https://aadrm.com/user_impersonation";

std::string Trim(const std::string& s) {
  size_t b = s.find_first_not_of(" \t\r\n");
  if (b == std::string::npos) return "";
  size_t e = s.find_last_not_of(" \t\r\n");
  return s.substr(b, e - b + 1);
}

std::string GetEnvOr(const char* name, const std::string& fallback) {
  const char* value = std::getenv(name);
  return (value != nullptr && value[0] != '\0') ? value : fallback;
}

bool LooksLikeRightsFailure(const std::string& what) {
  std::string lower = what;
  for (char& c : lower) c = static_cast<char>(tolower(c));
  return lower.find("right") != std::string::npos ||
      lower.find("access denied") != std::string::npos ||
      lower.find("access-denied") != std::string::npos ||
      lower.find("not authorized") != std::string::npos ||
      lower.find("forbidden") != std::string::npos ||
      lower.find("no permission") != std::string::npos;
}

std::string RunHelper(const std::string& helper,
                      const std::string& clientId,
                      const std::string& authority,
                      const std::string& scope) {
  // Single-quote each argument: helper paths and scopes are operator-controlled,
  // but quoting keeps a stray space from becoming a second command. stderr is
  // deliberately NOT merged: it inherits the terminal, so helper diagnostics
  // stay visible instead of becoming the "token".
  std::string cmd = "'" + helper + "' '" + clientId + "' '" + authority + "' '" + scope + "'";
  std::string out;
  char buf[512];
  FILE* raw = popen(cmd.c_str(), "r");
  if (raw == nullptr) throw std::runtime_error("popen(auth helper) failed");
  while (fgets(buf, sizeof(buf), raw) != nullptr) out += buf;
  // The exit status is load-bearing: a helper that tracebacked must fail the
  // decrypt loudly. Reading its last line as a token instead sends garbage to
  // RMS, which answers 401 — a real failure this exact shape once produced.
  // pclose returns the waitpid-encoded status, so decode it: reporting the
  // raw value would print 256 for an exit(1).
  int status = pclose(raw);
  int exitCode = (WIFEXITED(status) != 0) ? WEXITSTATUS(status) : -1;
  if (exitCode != 0) {
    throw std::runtime_error("auth helper exited with status " + std::to_string(exitCode));
  }
  // Token on stdout, diagnostics on stderr; the last non-empty line wins so a
  // chatty helper still yields its token.
  size_t end = out.find_last_not_of(" \t\r\n");
  if (end == std::string::npos) return "";
  size_t start = out.find_last_of('\n', end);
  std::string token = Trim(start == std::string::npos ? out : out.substr(start + 1));
  // RMS tokens are JWTs. Anything else is helper chatter that slipped onto
  // stdout — refuse it here rather than laundering it into a 401.
  size_t dots = 0;
  for (char c : token) dots += (c == '.');
  if (dots != 2 || token.size() < 32) {
    throw std::runtime_error("auth helper printed no JWT-shaped token on stdout");
  }
  return token;
}

class HelperAuth final : public mip::AuthDelegate {
 public:
  HelperAuth(std::string clientId, std::string defaultAuthority, std::string helper)
      : mClientId(std::move(clientId)),
        mAuthority(std::move(defaultAuthority)),
        mHelper(std::move(helper)) {}

  bool AcquireOAuth2Token(
      const mip::Identity& /*identity*/,
      const OAuth2Challenge& challenge,
      const std::shared_ptr<void>& /*context*/,
      OAuth2Token& token) override {
    std::string authority =
        challenge.GetAuthority().empty() ? mAuthority : challenge.GetAuthority();
    std::string res = challenge.GetResource();
    while (!res.empty() && res.back() == '/') res.pop_back();
    // MIP sends the tenant id as "scope" for RMS challenges; that is not a v2
    // scope (MSAL reads it as a Graph scope -> AADSTS650053). Map by resource.
    std::string scope = challenge.GetScope();
    bool scopeUsable = scope.find('/') != std::string::npos;
    if (res == kAadrmResource) {
      scope = kAadrmScope;
    } else if (!scopeUsable) {
      scope = res + "/.default";
    }
    // Authority/resource/scope name the challenge, not the credential, so they
    // are safe to log; the token itself is never printed.
    std::cerr << "[mip] auth authority=" << authority << " resource=" << challenge.GetResource()
              << " scope=" << scope << std::endl;
    std::string accessToken;
    try {
      accessToken = RunHelper(mHelper, mClientId, authority, scope);
    } catch (const std::exception& ex) {
      std::cerr << "[mip] auth helper failed: " << ex.what() << std::endl;
      return false;
    }
    if (accessToken.empty()) {
      std::cerr << "[mip] auth helper produced no token" << std::endl;
      return false;
    }
    token.SetAccessToken(accessToken);
    return true;
  }

 private:
  const std::string mClientId;
  const std::string mAuthority;
  const std::string mHelper;
};

class AcceptConsent final : public mip::ConsentDelegate {
 public:
  mip::Consent GetUserConsent(const std::string& /*url*/) override { return mip::Consent::Accept; }
};

class NoopProfileObserver final : public mip::FileProfile::Observer {
 public:
  void OnLoadSuccess(
      const std::shared_ptr<mip::FileProfile>& profile,
      const std::shared_ptr<void>& context) override {
    static_cast<std::promise<std::shared_ptr<mip::FileProfile>>*>(context.get())->set_value(profile);
  }
  void OnLoadFailure(
      const std::exception_ptr& error, const std::shared_ptr<void>& context) override {
    static_cast<std::promise<std::shared_ptr<mip::FileProfile>>*>(context.get())
        ->set_exception(error);
  }
  void OnAddEngineSuccess(
      const std::shared_ptr<mip::FileEngine>& engine,
      const std::shared_ptr<void>& context) override {
    static_cast<std::promise<std::shared_ptr<mip::FileEngine>>*>(context.get())->set_value(engine);
  }
  void OnAddEngineFailure(
      const std::exception_ptr& error, const std::shared_ptr<void>& context) override {
    static_cast<std::promise<std::shared_ptr<mip::FileEngine>>*>(context.get())
        ->set_exception(error);
  }
};

class NoopHandlerObserver final : public mip::FileHandler::Observer {
 public:
  void OnCreateFileHandlerSuccess(
      const std::shared_ptr<mip::FileHandler>& handler,
      const std::shared_ptr<void>& context) override {
    static_cast<std::promise<std::shared_ptr<mip::FileHandler>>*>(context.get())->set_value(handler);
  }
  void OnCreateFileHandlerFailure(
      const std::exception_ptr& error, const std::shared_ptr<void>& context) override {
    static_cast<std::promise<std::shared_ptr<mip::FileHandler>>*>(context.get())
        ->set_exception(error);
  }
  void OnCommitSuccess(bool committed, const std::shared_ptr<void>& context) override {
    static_cast<std::promise<bool>*>(context.get())->set_value(committed);
  }
  void OnCommitFailure(
      const std::exception_ptr& error, const std::shared_ptr<void>& context) override {
    static_cast<std::promise<bool>*>(context.get())->set_exception(error);
  }
};

class RestExecutionState final : public mip::FileExecutionState {};

} // namespace

int main(int argc, char** argv) {
  if (argc != 3) {
    std::cerr << "usage: mip-decrypt <input> <output>\n";
    return 2;
  }
  const std::string input = argv[1], output = argv[2];
  const std::string clientId = GetEnvOr("MIP_DECRYPT_CLIENT_ID", "");
  const std::string username = GetEnvOr("MIP_DECRYPT_USERNAME", "");
  const std::string helper = GetEnvOr("MIP_DECRYPT_AUTH_HELPER", "");
  if (clientId.empty() || username.empty() || helper.empty()) {
    std::cerr << "[mip] missing configuration: set MIP_DECRYPT_CLIENT_ID, "
                 "MIP_DECRYPT_USERNAME and MIP_DECRYPT_AUTH_HELPER\n";
    return 2;
  }
  const std::string authority = GetEnvOr("MIP_DECRYPT_AUTHORITY", kDefaultAuthority);
  std::string storage = GetEnvOr("MIP_DECRYPT_STORAGE", "");
  if (storage.empty()) {
    std::error_code ec;
    storage = (std::filesystem::temp_directory_path(ec) / "mip-decrypt").string();
    if (ec) {
      std::cerr << "[mip] cannot resolve a temp dir; set MIP_DECRYPT_STORAGE\n";
      return 2;
    }
  }

  try {
    mip::ApplicationInfo appInfo;
    appInfo.applicationId = clientId;
    appInfo.applicationName = "MipDecrypt";
    appInfo.applicationVersion = "1.0.0";
    auto mipConfig = std::make_shared<mip::MipConfiguration>(
        appInfo, storage, mip::LogLevel::Warning, false);
    auto mipContext = mip::MipContext::Create(mipConfig);

    auto authDelegate = std::make_shared<HelperAuth>(clientId, authority, helper);
    auto consent = std::make_shared<AcceptConsent>();
    mip::FileProfile::Settings profileSettings(
        mipContext, mip::CacheStorageType::InMemory, consent,
        std::make_shared<NoopProfileObserver>());
    auto profilePromise = std::make_shared<std::promise<std::shared_ptr<mip::FileProfile>>>();
    mip::FileProfile::LoadAsync(profileSettings, profilePromise);
    auto profile = profilePromise->get_future().get();

    mip::FileEngine::Settings engineSettings(
        mip::Identity(username), authDelegate, "" /*clientData*/, "en-US", false);
    engineSettings.SetProtectionOnlyEngine(true); // Skip label policy; RMS only.
    auto enginePromise = std::make_shared<std::promise<std::shared_ptr<mip::FileEngine>>>();
    profile->AddEngineAsync(engineSettings, enginePromise);
    auto engine = enginePromise->get_future().get();

    auto status = mip::FileHandler::GetFileStatus(input, mipContext);
    if (!status->IsProtected() && !status->ContainsProtectedObjects()) {
      // Not actually protected: copy through so exit 0 always means "output
      // is readable plaintext", and the caller never branches on protection.
      std::error_code ec;
      std::filesystem::copy_file(
          input, output, std::filesystem::copy_options::overwrite_existing, ec);
      if (ec) {
        std::cerr << "[mip] passthrough copy failed: " << ec.message() << "\n";
        return 1;
      }
      std::cerr << "[mip] not protected; copied through\n";
      return 0;
    }

    auto handlerPromise = std::make_shared<std::promise<std::shared_ptr<mip::FileHandler>>>();
    engine->CreateFileHandlerAsync(
        input, input, true /*auditDiscovery*/,
        std::make_shared<NoopHandlerObserver>(), handlerPromise,
        std::make_shared<RestExecutionState>());
    auto handler = handlerPromise->get_future().get();

    handler->RemoveProtection();
    if (!handler->IsModified()) {
      std::cerr << "[mip] RemoveProtection made no change\n";
      return 1;
    }
    auto commitPromise = std::make_shared<std::promise<bool>>();
    handler->CommitAsync(output, commitPromise);
    bool ok = commitPromise->get_future().get();
    std::cerr << (ok ? "[mip] WROTE " : "[mip] COMMIT FAILED ") << output << "\n";
    return ok ? 0 : 1;
  } catch (const std::exception& ex) {
    // The NO_RIGHTS marker is the launcher's branch signal; keep the keyword
    // stable and the human text free to change.
    if (LooksLikeRightsFailure(ex.what())) {
      std::cerr << "[mip] NO_RIGHTS: " << ex.what() << "\n";
    } else {
      std::cerr << "[mip] ERROR: " << ex.what() << "\n";
    }
    return 1;
  }
}
