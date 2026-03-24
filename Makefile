iAPP_NAME := KATS-Tools
VERSION := 1.0.0

BUILD_DIR := build

# -----------------------------
# macOS pkg
# -----------------------------
PKG_ID := se.example.kats-tools
PAYLOAD_DIR := $(BUILD_DIR)/payload
SCRIPTS_DIR := pkg/scripts
PKG_UNSIGNED := $(BUILD_DIR)/$(APP_NAME)-unsigned.pkg
PKG_SIGNED := $(BUILD_DIR)/$(APP_NAME)-$(VERSION).pkg

INSTALLER_CERT ?= Developer ID Installer: YOUR NAME (TEAMID)
NOTARY_PROFILE ?= KATS-NOTARY

# -----------------------------
# Windows installer (Inno Setup)
# -----------------------------
WIN_BUILD_DIR := $(BUILD_DIR)/windows
WIN_ISS := windows/KATS-Tools.iss
ISCC ?= ISCC.exe
WIN_INSTALLER_BASENAME := $(APP_NAME)-Setup-$(VERSION)

.PHONY: all clean mac-pkg mac-release-pkg windows-installer windows-installer-signed payload unsigned-pkg signed-pkg notarize

all: mac-pkg windows-installer

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
	productsign \
	  --sign "$(INSTALLER_CERT)" \
	  "$(PKG_UNSIGNED)" \
	  "$(PKG_SIGNED)"

notarize: signed-pkg
	xcrun notarytool submit "$(PKG_SIGNED)" \
	  --keychain-profile "$(NOTARY_PROFILE)" \
	  --wait
	xcrun stapler staple "$(PKG_SIGNED)"

mac-pkg: signed-pkg

mac-release-pkg: notarize

# ============================================================
# Windows
# ============================================================

windows-installer:
	mkdir -p "$(WIN_BUILD_DIR)"
	"$(ISCC)" /Qp /O"$(abspath $(WIN_BUILD_DIR))" /F"$(WIN_INSTALLER_BASENAME)" "$(WIN_ISS)"

# Placeholder om du senare vill signera med signtool i Windows CI
windows-installer-signed: windows-installer
	@echo "Unsigned Windows installer built at $(WIN_BUILD_DIR)/$(WIN_INSTALLER_BASENAME).exe"
	@echo "Add signtool here when you are ready to automate Windows code signing."

