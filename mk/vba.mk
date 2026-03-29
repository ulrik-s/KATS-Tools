VBA_MODULES := \
	src/KATSUtils.bas \
	src/RE.bas \
	src/MoGUser.bas \
	src/DebugSave.bas \
	src/DebugMail.bas \
	src/ModUpdate.bas \
	src/TagHandler.bas \
	src/Processor_KR.bas \
	src/Processor_YTTRANDE.bas \
	src/KATSMain.bas

VBA_STAGE_DIR := $(BUILD_DIR)/vba-src
VBA_MERGED := $(BUILD_DIR)/KATS-All.bas
VBA_IMPORT_README := $(VBA_STAGE_DIR)/IMPORT-ORDER.txt
VBA_MODULE_NAME ?= KATS_All

.PHONY: show-vba-order stage-vba merge-vba clean-vba

show-vba-order:
	@echo "VBA import order:"
	@for f in $(VBA_MODULES); do echo "  $$f"; done

stage-vba: $(VBA_MODULES)
	$(RM_RF) "$(VBA_STAGE_DIR)"
	$(MKDIR_P) "$(VBA_STAGE_DIR)"
	@i=1; \
	for f in $(VBA_MODULES); do \
		n=$$(printf '%02d' $$i); \
		base=$$(basename "$$f"); \
		cp "$$f" "$(VBA_STAGE_DIR)/$$n-$$base"; \
		i=$$((i+1)); \
	done
	@{ \
		echo "Import these files into VBA in this order:"; \
		echo ""; \
		i=1; \
		for f in $(VBA_MODULES); do \
			n=$$(printf '%02d' $$i); \
			base=$$(basename "$$f"); \
			echo "  $(VBA_STAGE_DIR)/$$n-$$base"; \
			i=$$((i+1)); \
		done; \
		echo ""; \
		echo "Recommended manual workflow:"; \
		echo "  1. Open KATS-Tools.dotm in Word VBA IDE"; \
		echo "  2. Remove old modules you want to replace"; \
		echo "  3. Import files in the order above"; \
		echo "  4. Save KATS-Tools.dotm"; \
	} > "$(VBA_IMPORT_README)"
	@echo "Staged VBA modules in $(VBA_STAGE_DIR)"
	@echo "See $(VBA_IMPORT_README)"

merge-vba: $(VBA_MODULES)
	$(MKDIR_P) "$(BUILD_DIR)"
	$(PYTHON) scripts/merge_vba.py \
		--output "$(VBA_MERGED)" \
		--module-name "$(VBA_MODULE_NAME)" \
		$(VBA_MODULES)
	@echo "Merged VBA module written to $(VBA_MERGED)"
	@echo "Warning: this flattens module boundaries and may expose Private-name collisions."

clean-vba:
	$(RM_RF) "$(VBA_STAGE_DIR)"
	$(RM_RF) "$(VBA_MERGED)"
