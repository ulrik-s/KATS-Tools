PKG_ID ?= se.qnyx.kats-tools
PAYLOAD_DIR := $(BUILD_DIR)/payload
SCRIPTS_DIR := pkg/scripts
PKG_UNSIGNED := $(BUILD_DIR)/$(APP_NAME)-unsigned.pkg
PKG_SIGNED := $(BUILD_DIR)/$(APP_NAME)-$(VERSION).pkg

INSTALLER_CERT ?=
NOTARY_PROFILE ?=

ifeq ($(UNAME_S),Darwin)
MAC_BASE_GROUP_DIR := $(HOME)/Library/Group Containers/UBF8T346G9.Office
MAC_WORD_STARTUP_LOCALIZED := $(MAC_BASE_GROUP_DIR)/User Content.localized/Startup.localized/Word
MAC_WORD_STARTUP_PLAIN := $(MAC_BASE_GROUP_DIR)/User Content/Startup/Word
MAC_WORD_STARTUP_DIR := $(shell if [ -d "$(MAC_WORD_STARTUP_LOCALIZED)" ]; then printf '%s' "$(MAC_WORD_STARTUP_LOCALIZED)"; elif [ -d "$(MAC_WORD_STARTUP_PLAIN)" ]; then printf '%s' "$(MAC_WORD_STARTUP_PLAIN)"; else printf '%s' "$(MAC_WORD_STARTUP_LOCALIZED)"; fi)
MAC_APP_SCRIPTS_DIR := $(HOME)/Library/Application Scripts/com.microsoft.Word
else
MAC_BASE_GROUP_DIR :=
MAC_WORD_STARTUP_LOCALIZED :=
MAC_WORD_STARTUP_PLAIN :=
MAC_WORD_STARTUP_DIR :=
MAC_APP_SCRIPTS_DIR :=
endif

.PHONY: check-postinstall payload unsigned-pkg signed-pkg notarize \
        mac-pkg mac-release-pkg install uninstall show-install-path

check-postinstall:
	@test -f "$(SCRIPTS_DIR)/postinstall" || (echo "Missing $(SCRIPTS_DIR)/postinstall" && exit 1)
	@file "$(SCRIPTS_DIR)/postinstall"
	@ls -l "$(SCRIPTS_DIR)/postinstall"

payload: build-dotm build-version-file
	$(RM_RF) "$(PAYLOAD_DIR)"
	$(MKDIR_P) "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)"
	cp "$(OUT_DOTM)" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Tools.dotm"
	cp "$(VERSION_FILE)" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Version.txt"
	cp "assets/KATSUpdater.applescript" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.applescript"
	cp "assets/KATSUpdater.sh" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.sh"
	cp "assets/KATSMail.applescript" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSMail.applescript"
	cp "assets/KATSFileOps.applescript" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSFileOps.applescript"
	cp "assets/KATSUpdater.bat" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.bat"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Tools.dotm"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Version.txt"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.applescript"
	chmod 755 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.sh"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSMail.applescript"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSFileOps.applescript"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.bat"

unsigned-pkg: check-postinstall payload
	chmod +x "$(SCRIPTS_DIR)/postinstall"
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
release-all: mac-release-pkg

show-install-path:
	@echo "$(MAC_WORD_STARTUP_DIR)"

install: build-dotm build-version-file
	@if [ -z "$(MAC_WORD_STARTUP_DIR)" ]; then echo "Could not determine Word Startup folder."; exit 1; fi
	$(MKDIR_P) "$(MAC_WORD_STARTUP_DIR)"
	$(MKDIR_P) "$(MAC_APP_SCRIPTS_DIR)"
	cp "$(OUT_DOTM)" "$(MAC_WORD_STARTUP_DIR)/KATS-Tools.dotm"
	cp "$(VERSION_FILE)" "$(MAC_WORD_STARTUP_DIR)/KATS-Version.txt"
	cp "assets/KATSUpdater.applescript" "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.applescript"
	cp "assets/KATSUpdater.sh" "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.sh"
	cp "assets/KATSMail.applescript" "$(MAC_APP_SCRIPTS_DIR)/KATSMail.applescript"
	cp "assets/KATSFileOps.applescript" "$(MAC_APP_SCRIPTS_DIR)/KATSFileOps.applescript"
	chmod 644 "$(MAC_WORD_STARTUP_DIR)/KATS-Tools.dotm"
	chmod 644 "$(MAC_WORD_STARTUP_DIR)/KATS-Version.txt"
	chmod 644 "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.applescript"
	chmod 755 "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.sh"
	chmod 644 "$(MAC_APP_SCRIPTS_DIR)/KATSMail.applescript"
	chmod 644 "$(MAC_APP_SCRIPTS_DIR)/KATSFileOps.applescript"
	@echo "Installed to $(MAC_WORD_STARTUP_DIR)"
	@echo "AppleScripts installed to $(MAC_APP_SCRIPTS_DIR)"
	@ls -l "$(MAC_WORD_STARTUP_DIR)/KATS-Tools.dotm" "$(MAC_WORD_STARTUP_DIR)/KATS-Version.txt" "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.applescript" "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.sh" "$(MAC_APP_SCRIPTS_DIR)/KATSMail.applescript" "$(MAC_APP_SCRIPTS_DIR)/KATSFileOps.applescript"
	@echo "Restart Word completely."

uninstall:
	@if [ -z "$(MAC_WORD_STARTUP_DIR)" ]; then echo "Could not determine Word Startup folder."; exit 1; fi
	rm -f "$(MAC_WORD_STARTUP_DIR)/KATS-Tools.dotm"
	rm -f "$(MAC_WORD_STARTUP_DIR)/KATS-Version.txt"
	rm -f "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.applescript"
	rm -f "$(MAC_APP_SCRIPTS_DIR)/KATSUpdater.sh"
	rm -f "$(MAC_APP_SCRIPTS_DIR)/KATSMail.applescript"
	rm -f "$(MAC_APP_SCRIPTS_DIR)/KATSFileOps.applescript"
	@echo "Removed."
	@echo "Restart Word completely."
