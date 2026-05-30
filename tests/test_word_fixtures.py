#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Unittest entrypoint for Word fixture smoke checks."""

from __future__ import annotations

import unittest

from tests.fixtures import check_word_samples


class WordFixtureSmokeTest(unittest.TestCase):
    def test_generated_samples_emit_expected_diagnostics(self) -> None:
        self.assertEqual(check_word_samples.main(), 0)


if __name__ == "__main__":
    unittest.main()
