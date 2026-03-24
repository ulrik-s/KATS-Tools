.PHONY: explode-dotm inject-customui patch-ribbon-rels patch-content-types \
        build-dotm rebuild-dotm show-workdir

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

patch-ribbon-rels: inject-customui
	@test -f "$(WORK_DIR)/_rels/.rels" || (echo "Missing $(WORK_DIR)/_rels/.rels" && exit 1)
	$(PYTHON) scripts/patch_ribbon_rels.py "$(WORK_DIR)/_rels/.rels"

patch-content-types: patch-ribbon-rels
	@test -f "$(WORK_DIR)/[Content_Types].xml" || (echo "Missing $(WORK_DIR)/[Content_Types].xml" && exit 1)
	$(PYTHON) scripts/patch_content_types.py "$(WORK_DIR)/[Content_Types].xml"

build-dotm: patch-content-types
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

