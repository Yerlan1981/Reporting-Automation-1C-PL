"""Автоматические проверки ключевых мест программы. Запуск: python -m pytest"""
import pandas as pd
import pytest

from Reporting_Automation import (
    apply_mapping,
    apply_shares,
    check_totals,
    get_period,
    get_unrecognized,
    validate_settings,
)


def test_period_is_read_from_header():
    df = pd.DataFrame({0: ["Оборотно-сальдовая ведомость по счету 6010  за Март 2025 г."]})
    assert get_period(df) == "2025-03"


def test_shares_do_not_lose_money():
    # одна статья СУГ на 1000, в справочнике она стоит двумя строками (газ и логистика)
    merged = pd.DataFrame(
        {
            "pl_line": ["ТМЗ - основное производство", "ТМЗ - логистические расходы"],
            "amount": [1000.0, 1000.0],
            "sign": [1, 1],
        }
    )
    settings = {"gas_share": 0.7, "logistics_share": 0.3}

    result = apply_shares(merged, settings)

    assert result["final_amount"].tolist() == pytest.approx([700.0, 300.0])
    assert result["final_amount"].sum() == pytest.approx(1000.0)


def test_shares_must_sum_to_100_percent():
    with pytest.raises(ValueError):
        validate_settings({"gas_share": 0.6, "logistics_share": 0.6})

    validate_settings({"gas_share": 0.6, "logistics_share": 0.4})  # не должно падать


def test_unknown_article_goes_to_fallback_line():
    osv = pd.DataFrame(
        {
            "account": ["7110"],
            "article": ["Выдуманная статья, которой нет в справочнике"],
            "amount": [500.0],
        }
    )

    merged = apply_mapping(osv)

    assert merged["pl_line"].iloc[0] == "Прочее"  # запасная строка для счёта 7110
    assert len(get_unrecognized(merged)) == 1  # и она попала в список нераспознанных


def test_control_detects_lost_amount():
    osv = pd.DataFrame({"account": ["7110", "7110"], "amount": [100.0, 50.0]})

    # обе статьи попали в отчёт: разницы нет
    ok = pd.DataFrame(
        {"account": ["7110", "7110"], "amount": [100.0, 50.0], "share": [1, 1], "pl_line": ["A", "B"]}
    )
    assert check_totals(osv, ok)["diff"].tolist() == [0.0]

    # вторая статья никуда не попала: контроль должен показать потерю 50
    lost = ok.copy()
    lost["pl_line"] = ["A", None]
    assert check_totals(osv, lost)["diff"].tolist() == [50.0]


def test_kpn_account_rule_needs_no_mapping():
    # счёт 7710: любая статья идёт в строку КПН и не считается нераспознанной
    osv = pd.DataFrame({"account": ["7710"], "article": ["Любое название статьи"], "amount": [100.0]})

    merged = apply_mapping(osv)

    assert merged["pl_line"].iloc[0] == "КПН"
    assert get_unrecognized(merged).empty
    