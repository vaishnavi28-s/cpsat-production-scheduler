"""
Unit tests for the print job optimizer.
Run with: pytest tests/
"""

import pytest
from src.models import PrintJob
from src.finishing_rules import (
    type_label, finishing_conflict, special_ink_set,
    uv_overall, uv_spot, has_lamination, has_foil, has_emboss
)
from src.optimizer import solve_group, is_28x20_sheet


def make_job(**kwargs) -> PrintJob:
    defaults = {
        "JOB": "99999",
        "PRESS_LOCATION": "Martinsburg - BVG",
        "SEND_TO_LOCATION": "Martinsburg - BVG",
        "PRODUCTTYPE": "Cover",
        "PAPER": "100# Gloss Text 28 x 20",
        "FINISHTYPE": "",
        "FINISHINGOP": "",
        "DELIVERYDATE": "2026-06-01",
        "INKSS1": "C, M, Y, K",
        "INKSS2": "",
        "QUANTITYORDERED": 5000,
        "PAGES": 128,
    }
    defaults.update(kwargs)
    return PrintJob(**defaults)


class TestTypeLabel:
    def test_4_0(self):
        j = make_job(INKSS1="C, M, Y, K", INKSS2="")
        assert type_label(j) == "4/0"

    def test_4_1(self):
        j = make_job(INKSS1="C, M, Y, K", INKSS2="K")
        assert type_label(j) == "4/1"

    def test_5_0(self):
        j = make_job(INKSS1="C, M, Y, K, PMS 485", INKSS2="")
        assert type_label(j) == "5/0"

    def test_4_4(self):
        j = make_job(INKSS1="C, M, Y, K", INKSS2="C, M, Y, K")
        assert type_label(j) == "4/4"


class TestFinishingConflict:
    def test_uv_overall_with_lamination_conflicts(self):
        a = make_job(JOB="A1", FINISHTYPE="UV Overall", FINISHINGOP="")
        b = make_job(JOB="B1", FINISHTYPE="Lamination", FINISHINGOP="Gloss Polypropylene")
        assert finishing_conflict(a, b) is True

    def test_uv_overall_with_spot_only_conflicts(self):
        a = make_job(JOB="A2", FINISHTYPE="UV Overall", FINISHINGOP="")
        b = make_job(JOB="B2", FINISHTYPE="UV Spot", FINISHINGOP="Spot UV")
        assert finishing_conflict(a, b) is True

    def test_same_lamination_compatible(self):
        a = make_job(JOB="A3", FINISHTYPE="Lamination", FINISHINGOP="Gloss Polypropylene")
        b = make_job(JOB="B3", FINISHTYPE="Lamination", FINISHINGOP="Gloss Polypropylene")
        assert finishing_conflict(a, b) is False

    def test_gloss_vs_matte_lam_conflicts(self):
        a = make_job(JOB="A4", FINISHTYPE="Lamination", FINISHINGOP="Gloss Polypropylene")
        b = make_job(JOB="B4", FINISHTYPE="Lamination", FINISHINGOP="Matte Polypropylene")
        assert finishing_conflict(a, b) is True

    def test_press_varnish_with_lam_conflicts(self):
        a = make_job(JOB="A5", FINISHTYPE="Press Varnish", FINISHINGOP="Aqueous")
        b = make_job(JOB="B5", FINISHTYPE="Lamination", FINISHINGOP="Matte Polypropylene")
        assert finishing_conflict(a, b) is True

    def test_no_conflict_plain_jobs(self):
        a = make_job(JOB="A6")
        b = make_job(JOB="B6")
        assert finishing_conflict(a, b) is False


class TestSpecialInkSet:
    def test_process_inks_excluded(self):
        j = make_job(INKSS1="C, M, Y, K", INKSS2="")
        assert special_ink_set(j) == set()

    def test_special_ink_detected(self):
        j = make_job(INKSS1="C, M, Y, K, PMS 485", INKSS2="")
        s = special_ink_set(j)
        assert len(s) == 1

    def test_black_is_special(self):
        j = make_job(INKSS1="C, M, Y, K, Black", INKSS2="")
        s = special_ink_set(j)
        assert len(s) == 1


class TestSheetSize:
    def test_28x20_detected(self):
        j = make_job(PAPER="100# Gloss Text 28 x 20")
        assert is_28x20_sheet(j) is True

    def test_28x20_half_detected(self):
        j = make_job(PAPER="80# Gloss Text 28 x 20 1/2")
        assert is_28x20_sheet(j) is True

    def test_25x38_not_28x20(self):
        j = make_job(PAPER="100# Gloss Text 25 x 38")
        assert is_28x20_sheet(j) is False


class TestOptimizer:
    def test_single_job_returns_singleton(self):
        jobs = [make_job(JOB="J1", PRODUCTTYPE="Cover")]
        result = solve_group(jobs)
        assert len(result) == 1
        assert result[0] == ["J1"]

    def test_compatible_covers_combined(self):
        jobs = [
            make_job(JOB="C1", PRODUCTTYPE="Cover", QUANTITYORDERED=5000),
            make_job(JOB="C2", PRODUCTTYPE="Cover", QUANTITYORDERED=5000),
        ]
        result = solve_group(jobs)
        assert len(result) == 1
        assert set(result[0]) == {"C1", "C2"}

    def test_qty_conflict_prevents_combination(self):
        jobs = [
            make_job(JOB="Q1", PRODUCTTYPE="Cover", QUANTITYORDERED=1000),
            make_job(JOB="Q2", PRODUCTTYPE="Cover", QUANTITYORDERED=10000),
        ]
        result = solve_group(jobs)
        assert len(result) == 2

    def test_jacket_capacity_is_2(self):
        jobs = [
            make_job(JOB=f"J{i}", PRODUCTTYPE="Jacket", QUANTITYORDERED=3000)
            for i in range(4)
        ]
        result = solve_group(jobs)
        for combo in result:
            assert len(combo) <= 2

    def test_finish_conflict_prevents_combination(self):
        jobs = [
            make_job(JOB="F1", PRODUCTTYPE="Cover",
                     FINISHTYPE="UV Overall", FINISHINGOP="", QUANTITYORDERED=3000),
            make_job(JOB="F2", PRODUCTTYPE="Cover",
                     FINISHTYPE="Lamination", FINISHINGOP="Gloss Polypropylene", QUANTITYORDERED=3000),
        ]
        result = solve_group(jobs)
        assert len(result) == 2
