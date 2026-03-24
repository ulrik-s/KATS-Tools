PKG_ID ?= se.qnyx.kats-tools
PAYLOAD_DIR := $(BUILD_DIR)/payload
SCRIPTS_DIR := pkg/scripts
PKG_UNSIGNED := $(BUILD_DIR)/$(APP_NAME)-unsigned.pkg
PKG_SIGNED := $(BUILD_DIR)/$(APP_NAME)-$(VERSION).pkg

INSTALLER_CERT ?=
NOTARY_PROFILE ?=

.PHONY: check-postinstall payload unsigned-pkg signed-pkg notarize mac-pkg mac-release-pkg

check-postinstall:
	@test -f "$(SCRIPTS_DIR)/postinstall" || (echo "Missing $(SCRIPTS_DIR)/postinstall" && exit 1)
	@file "$(SCRIPTS_DIR)/postinstall"
	@ls -l "$(SCRIPTS_DIR)/postinstall"

payload: build-dotm
	$(RM_RF) "$(PAYLOAD_DIR)"
	$(MKDIR_P) "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)"
	cp "$(OUT_DOTM)" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Tools.dotm"
	cp "assets/KATSUpdater.applescript" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.applescript"
	cp "assets/KATSUpdater.bat" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.bat"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Tools.dotm"
	chmod 644 "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATSUpdater.applescript"
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

