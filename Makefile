APP_NAME ?= KATS-Tools
VERSION ?= 1.0.0

BUILD_DIR := build
WORK_ROOT := work
WORK_DIR := $(WORK_ROOT)/$(APP_NAME)

UNAME_S := $(shell uname -s 2>/dev/null)
IS_WINDOWS := $(filter Windows_NT,$(OS))

# ============================================================
# DOTM rebuild from seed + customUI.xml
# ============================================================

SEED_DOTM ?= assets/$(APP_NAME).dotm
CUSTOMUI_XML ?= ribbon/customUI.xml
OUT_DOTM := $(BUILD_DIR)/$(APP_NAME).dotm

UNZIP ?= unzip
ZIP ?= zip
MKDIR_P ?= mkdir -p
RM_RF ?= rm -rf
CP ?= cp
FIND ?= find
POWERSHELL ?= powershell -NoProfile -ExecutionPolicy Bypass

# ============================================================
# macOS pkg
# ============================================================

PKG_ID ?= se.qnyx.kats-tools
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
WIN_DOTM_FOR_INSTALLER := $(WIN_BUILD_DIR)/$(APP_NAME).dotm

.PHONY: all clean clean-dotm clean-build \
        explode-dotm inject-customui build-dotm rebuild-dotm show-workdir \
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
	@echo "  make build-dotm"
endif

release-all: mac-release-pkg windows-installer

clean: clean-dotm clean-build

clean-dotm:
	$(RM_RF) "$(WORK_DIR)"

clean-build:
	$(RM_RF) "$(BUILD_DIR)"

# ============================================================
# DOTM rebuild
# ============================================================

explode-dotm: clean-dotm
	@test -f "$(SEED_DOTM)" || (echo "Missing seed DOTM: $(SEED_DOTM)" && exit 1)
	$(MKDIR_P) "$(WORK_DIR)"
	$(UNZIP) -q "$(SEED_DOTM)" -d "$(WORK_DIR)"
	@echo "Exploded $(SEED_DOTM) -> $(WORK_DIR)"

inject-customui: explode-dotm
	@test -f "$(CUSTOMUI_XML)" || (echo "Missing customUI XML: $(CUSTOMUI_XML)" && exit 1)
	$(MKDIR_P) "$(WORK_DIR)/customUI"
	$(CP) "$(CUSTOMUI_XML)" "$(WORK_DIR)/customUI/customUI.xml"
	@echo "Injected $(CUSTOMUI_XML) -> $(WORK_DIR)/customUI/customUI.xml"

build-dotm: inject-customui
	@test -f "$(WORK_DIR)/[Content_Types].xml" || (echo "Missing [Content_Types].xml in $(WORK_DIR)" && exit 1)
	@test -f "$(WORK_DIR)/customUI/customUI.xml" || (echo "Missing customUI/customUI.xml in $(WORK_DIR)" && exit 1)
	$(MKDIR_P) "$(BUILD_DIR)"
	$(RM_RF) "$(OUT_DOTM)"
	$(FIND) "$(WORK_DIR)" -name ".DS_Store" -delete
ifeq ($(IS_WINDOWS),Windows_NT)
	$(POWERSHELL) -Command "Compress-Archive -Path '$(WORK_DIR)\\*' -DestinationPath '$(OUT_DOTM)' -Force"
else
	cd "$(WORK_DIR)" && $(ZIP) -X -q -r "$(abspath $(OUT_DOTM))" .
endif
	@echo "Built $(OUT_DOTM)"

rebuild-dotm: build-dotm

show-workdir:
	@test -d "$(WORK_DIR)" || (echo "Missing work dir: $(WORK_DIR)" && exit 1)
	@find "$(WORK_DIR)" | sort

# ============================================================
# macOS
# ============================================================

payload: build-dotm
	$(RM_RF) "$(PAYLOAD_DIR)"
	$(MKDIR_P) "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)"
	cp "$(OUT_DOTM)" "$(PAYLOAD_DIR)/Library/Application Support/$(APP_NAME)/KATS-Tools.dotm"
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

windows-installer: build-dotm
	mkdir -p "$(WIN_BUILD_DIR)"
	cp "$(OUT_DOTM)" "$(WIN_DOTM_FOR_INSTALLER)"
ifeq ($(UNAME_S),Darwin)
	@echo "Skipping Windows installer build on macOS."
	@echo "Rebuilt .dotm is ready at: $(OUT_DOTM)"
	@echo "Build Windows installer in CI or on a Windows machine."
else
	"$(ISCC)" /Qp /O"$(abspath $(WIN_BUILD_DIR))" /F"$(WIN_INSTALLER_BASENAME)" /DMyAppVersion=$(VERSION) "$(WIN_ISS)"
endif

windows-installer-signed: windows-installer
	@echo "Unsigned Windows installer built at:"
	@echo "  $(WIN_BUILD_DIR)/$(WIN_INSTALLER_BASENAME).exe"
	@echo "Add signtool in CI or a separate target when signing is ready."

