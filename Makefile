PYTHON ?= .venv/bin/python
PYCACHE ?= /private/tmp/gongwen-pycache

.PHONY: check smoke-word render-word render-word-required evals coverage-matrix golden-word skill-validate test compile

check: compile skill-validate test smoke-word evals coverage-matrix golden-word

smoke-word:
	$(PYTHON) tests/fixtures/check_word_samples.py

render-word:
	$(PYTHON) tests/fixtures/render_word_samples.py

render-word-required:
	$(PYTHON) tests/fixtures/render_word_samples.py --require-renderer

evals:
	$(PYTHON) scripts/run_evals.py

coverage-matrix:
	$(PYTHON) tests/fixtures/check_coverage_matrix.py

golden-word:
	$(PYTHON) tests/fixtures/check_golden_xml.py

skill-validate:
	PYTHONPATH=$(CURDIR)/scripts/quick_validate_yaml_shim $(PYTHON) /Users/abi/.codex/skills/.system/skill-creator/scripts/quick_validate.py skills/format-xingwen-word

test:
	$(PYTHON) -m unittest tests.test_word_fixtures

compile:
	PYTHONPYCACHEPREFIX=$(PYCACHE) $(PYTHON) -m py_compile \
		scripts/run_evals.py \
		scripts/quick_validate_yaml_shim/yaml.py \
		skills/format-xingwen-word/scripts/format_document.py \
		skills/format-xingwen-word/scripts/xingwen_word/__init__.py \
		skills/format-xingwen-word/scripts/xingwen_word/formatting.py \
		skills/format-xingwen-word/scripts/xingwen_word/presets.py \
		skills/format-xingwen-word/scripts/xingwen_word/reports.py \
		skills/format-xingwen-word/scripts/xingwen_word/roles.py \
		tests/fixtures/generate_word_samples.py \
		tests/fixtures/check_word_samples.py \
		tests/fixtures/render_word_samples.py \
		tests/fixtures/check_coverage_matrix.py \
		tests/fixtures/check_golden_xml.py \
		tests/test_word_fixtures.py
