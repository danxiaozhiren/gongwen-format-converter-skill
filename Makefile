PYTHON ?= .venv/bin/python
PYCACHE ?= /private/tmp/gongwen-pycache

.PHONY: smoke-word render-word render-word-required test compile

smoke-word:
	$(PYTHON) tests/fixtures/check_word_samples.py

render-word:
	$(PYTHON) tests/fixtures/render_word_samples.py

render-word-required:
	$(PYTHON) tests/fixtures/render_word_samples.py --require-renderer

test:
	$(PYTHON) -m unittest tests.test_word_fixtures

compile:
	PYTHONPYCACHEPREFIX=$(PYCACHE) $(PYTHON) -m py_compile \
		skills/gongwen-format-converter/scripts/format_document.py \
		tests/fixtures/generate_word_samples.py \
		tests/fixtures/check_word_samples.py \
		tests/fixtures/render_word_samples.py \
		tests/test_word_fixtures.py
