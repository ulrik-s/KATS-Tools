BUILD_DIR := build
WORK_ROOT := work
WORK_DIR := $(WORK_ROOT)/$(APP_NAME)

UNAME_S := $(shell uname -s 2>/dev/null)
IS_WINDOWS := $(filter Windows_NT,$(OS))

ifeq ($(IS_WINDOWS),Windows_NT)
PYTHON ?= python
else
PYTHON ?= python3
endif

UNZIP ?= unzip
ZIP ?= zip
MKDIR_P ?= mkdir -p
RM_RF ?= rm -rf
CP ?= cp
FIND ?= find
POWERSHELL ?= powershell -NoProfile -ExecutionPolicy Bypass

SEED_DOTM ?= assets/$(APP_NAME).dotm
CUSTOMUI_XML ?= ribbon/customUI.xml
OUT_DOTM := $(BUILD_DIR)/$(APP_NAME).dotm

.PHONY: all clean clean-dotm clean-build release-all

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
	@echo "  make install"
endif

clean: clean-dotm clean-build

clean-dotm:
	$(RM_RF) "$(WORK_DIR)"

clean-build:
	$(RM_RF) "$(BUILD_DIR)"

