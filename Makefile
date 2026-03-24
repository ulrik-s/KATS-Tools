APP_NAME ?= KATS-Tools
VERSION ?= 1.0.0

BUILD_DIR := build

UNAME_S := $(shell uname -s 2>/dev/null)
IS_WINDOWS := $(filter Windows_NT,$(OS))

# ============================================================
# macOS pkg
# ============================================================

PKG_ID ?= se.example.kats-tools
PAYLOAD_DIR := $(BUILD_DIR)/payload
SCRIPTS_DIR := pkg/scripts
PKG_UNSIGNED := $(BUILD_DIR)/$(APP_NAME)-unsigned.pkg
PKG_SIGNED := $(BUILD_DIR)/$(APP_NAME)-$(VERSION).pkg

INSTALLER_CERT ?=
NOTARY_PROFILE ?=

# ============================================================
# Windows installer (Inno Setup)
# ============================================================

WIN_BUILD_DIR := $(BUILD_DIR)/windows
WIN_ISS := windows/KATS-Tools.iss
ISCC ?= ISCC.exe
WIN_INSTALLER_BASENAME := $(APP_NAME)-Setup-$(VERSION)

.PHONY: all clean \
        payload unsigned-pkg signed-pkg notarize mac-pkg mac-release-pkg \
        windows-installer windows-installer-signed \
        release-all

# ------------------------------------------------------------
# Native default target
# ------------------------------------------------------------

ifeq ($(UNAME_S),Darwin)
all: mac-pkg
else ifeq ($(IS_WINDOWS),Windows_NT)
all: windows-installer
else
all:
	@echo "Unknown host OS. Use one of:"
	@echo "  make mac-pkg"
	@echo "  make windows-installer"
endif

# Build both, mainly useful in CI where each platform runs its own job
release-all: mac-release-pkg windows-installer

clean:
	rm -rf "$(BUILD_DIR)"

# ============================================================
# macOS
# ============================================================

payload: clean
	mkdir -p "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)"
	cp "assets/KATS-Tools.dotm" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Tools.dotm"
	cp "assets/KATSUpdater.applescript" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.applescript"
	cp "assets/KATSUpdater.bat" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.bat"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Tools.dotm"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.applescript"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.bat"

unsigned-pkg: payload
	pkgbuild \
	  --root "$(PAYLOAD_DIR)" \
	  --identifier "$(PKG_ID)" \
	  --version "$(VERSION)" \
	  --install-location "/" \
	  --scripts "$(SCRIPTS_DIR)" \
	  "$(PKG_UNSIGNED)"

signed-pkg: unsigned-pkg
ifeq ($(strip $(INSTALLER_CERT)),)
	@echo "No INSTALLER_CERT provided. Keeping unsigned pkg:"
	@echo "  $(PKG_UNSIGNED)"
else
	productsign \
	  --sign "$(INSTALLER_CERT)" \
	  "$(PKG_UNSIGNED)" \
	  "$(PKG_SIGNED)"
	@echo "Signed pkg built:"
	@echo "  $(PKG_SIGNED)"
endif

notarize: signed-pkg
ifeq ($(strip $(INSTALLER_CERT)),)
	@echo "Skipping notarization because pkg is unsigned."
else ifeq ($(strip $(NOTARY_PROFILE)),)
	@echo "Skipping notarization because NOTARY_PROFILE is not set."
else
	xcrun notarytool submit "$(PKG_SIGNED)" \
	  --keychain-profile "$(NOTARY_PROFILE)" \
	  --wait
	xcrun stapler staple "$(PKG_SIGNED)"
	@echo "Notarized pkg:"
	@echo "  $(PKG_SIGNED)"
endif

mac-pkg: signed-pkg

mac-release-pkg: notarize

# ============================================================
# Windows
# ============================================================

windows-installer:
	mkdir -p "$(WIN_BUILD_DIR)"
ifeq ($(UNAME_S),Darwin)
	@echo "Skipping Windows installer build on macOS."
	@echo "Build Windows installer in CI or on a Windows machine."
else
	"$(ISCC)" /Qp /O"$(abspath $(WIN_BUILD_DIR))" /F"$(WIN_INSTALLER_BASENAME)" "$(WIN_ISS)"
endif

windows-installer-signed: windows-installer
	@echo "Unsigned Windows installer built at:"
	@echo "  $(WIN_BUILD_DIR)/$(WIN_INSTALLER_BASENAME).exe"
	@echo "Add signtool in CI or a separate target when signing is ready."
