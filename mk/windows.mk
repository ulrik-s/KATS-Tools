WIN_BUILD_DIR := $(BUILD_DIR)/windows
WIN_ISS := windows/KATS-Tools.iss
ISCC ?= ISCC.exe
WIN_INSTALLER_BASENAME := $(APP_NAME)-Setup-$(VERSION)
WIN_DOTM_FOR_INSTALLER := $(WIN_BUILD_DIR)/$(APP_NAME).dotm
WIN_VERSION_FOR_INSTALLER := $(WIN_BUILD_DIR)/KATS-Version.txt

WIN_BUILD_DIR_WIN := $(subst /,\,$(abspath $(WIN_BUILD_DIR)))
WIN_ISS_WIN := $(subst /,\,$(WIN_ISS))
WIN_ISCC_WIN := $(subst /,\,$(ISCC))

.PHONY: windows-installer windows-installer-signed

windows-installer: build-dotm build-version-file
	mkdir -p "$(WIN_BUILD_DIR)"
	cp "$(OUT_DOTM)" "$(WIN_DOTM_FOR_INSTALLER)"
	cp "$(VERSION_FILE)" "$(WIN_VERSION_FOR_INSTALLER)"
ifeq ($(UNAME_S),Darwin)
	@echo "Skipping Windows installer build on macOS."
	@echo "Rebuilt .dotm is ready at: $(OUT_DOTM)"
	@echo "Version file is ready at: $(VERSION_FILE)"
	@echo "Build Windows installer in CI or on a Windows machine."
else
	$(POWERSHELL) -Command "& '$(WIN_ISCC_WIN)' '/Qp' '/O$(WIN_BUILD_DIR_WIN)' '/F$(WIN_INSTALLER_BASENAME)' '/DMyAppVersion=$(VERSION)' '$(WIN_ISS_WIN)'"
endif

windows-installer-signed: windows-installer
	@echo "Unsigned Windows installer built at:"
	@echo "  $(WIN_BUILD_DIR)/$(WIN_INSTALLER_BASENAME).exe"
	@echo "Add signtool in CI or a separate target when signing is ready."

release-all: windows-installer
