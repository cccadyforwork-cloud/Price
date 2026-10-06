#!/usr/bin/env python3
"""按当前工作台规则重算所有快速定价历史记录。"""

import argparse
import json
import math
import shutil
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "pricing_config.json"
HISTORY_PATH = ROOT / "config" / "quick_pricing_history.json"
BACKUP_DIR = ROOT / "config" / "backups"

SHIPPING_TIERS = {
    "under4": {"label": "≤4oz"},
    "4to8": {"label": "4–8oz", "fee": 1.77},
    "8to12": {"label": "8–12oz", "fee": 2.60},
    "12to16": {"label": "12–16oz", "fee": 3.22},
    "1to1_25": {"label": "1–1.25lb", "fee": 3.72},
}


def read_json(path):
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path, data):
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def round_up_cents(value):
    return math.ceil((float(value or 0) - 1e-9) * 100) / 100


def tier_from_record(record):
    tier_id = record.get("tierId")
    if tier_id in SHIPPING_TIERS:
        return tier_id, SHIPPING_TIERS[tier_id]
    fee = float(record.get("shippingFee") or 0)
    mappings = [
        ({0.50, 0.88, 0.99}, "under4"),
        ({1.77, 2.05}, "4to8"),
        ({2.60, 2.84}, "8to12"),
        ({3.22, 3.48}, "12to16"),
        ({3.72, 4.14}, "1to1_25"),
    ]
    for candidates, candidate_id in mappings:
        if any(abs(fee - candidate) < 0.01 for candidate in candidates):
            return candidate_id, SHIPPING_TIERS[candidate_id]
    return "under4", SHIPPING_TIERS["under4"]


def fee_for_tier(tier_id, tier, price):
    if tier_id == "under4":
        return 0.50 if price <= 3 else 0.88
    return float(tier["fee"])


def factors(config, return_rate):
    referral = float(config.get("referral_fee_rate", 0.15))
    return_rate = float(return_rate)
    return {
        "currency_rate": float(config.get("currency_rate_rmb_to_usd", 7.2)),
        "return_rate": return_rate,
        "first_leg": float(config.get("first_leg_shipping_usd", 0.3)),
        "disposal": float(config.get("disposal_fee_usd", 0.25)),
        "factor": (1 - return_rate) * (1 - referral) - return_rate * referral * referral,
    }


def calculate_in_regime(
    cost_rmb, tier_id, tier, margin, config, return_rate, reference_price,
    shipping_fee_override=None,
):
    values = factors(config, return_rate)
    cost_usd = float(cost_rmb or 0) / values["currency_rate"]
    shipping_fee = (
        float(shipping_fee_override)
        if shipping_fee_override not in (None, "")
        else fee_for_tier(tier_id, tier, reference_price)
    )
    fixed_cost = cost_usd + values["first_leg"] + shipping_fee + values["disposal"] * values["return_rate"]
    raw_price = fixed_cost / max(0.1, values["factor"] - margin)
    price = round_up_cents(raw_price)
    profit_usd = price * values["factor"] - fixed_cost
    return {
        "price": price,
        "raw_price": raw_price,
        "shipping_fee": shipping_fee,
        "cost_usd": cost_usd,
        "fixed_cost": fixed_cost,
        "profit_rmb": profit_usd * values["currency_rate"],
        "return_rate": values["return_rate"],
    }


def calculate_price(cost_rmb, tier_id, tier, margin, config, shipping_fee_override=None):
    threshold = float(config.get("return_rate_price_threshold_usd", 3.0))
    configured_rate = float(config.get("return_rate", 0.05))
    without_returns = calculate_in_regime(
        cost_rmb, tier_id, tier, margin, config, 0.0, threshold,
        shipping_fee_override,
    )
    if without_returns["price"] <= threshold:
        return without_returns

    with_returns = calculate_in_regime(
        cost_rmb, tier_id, tier, margin, config, configured_rate, threshold + 0.01,
        shipping_fee_override,
    )
    price = max(round_up_cents(threshold + 0.01), with_returns["price"])
    return_factors = factors(config, configured_rate)
    shipping_fee = (
        float(shipping_fee_override)
        if shipping_fee_override not in (None, "")
        else fee_for_tier(tier_id, tier, price)
    )
    fixed_cost = with_returns["cost_usd"] + return_factors["first_leg"] + shipping_fee + return_factors["disposal"] * return_factors["return_rate"]
    profit_usd = price * return_factors["factor"] - fixed_cost
    return {
        **with_returns,
        "price": price,
        "shipping_fee": shipping_fee,
        "fixed_cost": fixed_cost,
        "profit_rmb": profit_usd * return_factors["currency_rate"],
        "return_rate": configured_rate,
    }


def recalculate_row(row, tier_id, tier, config):
    previous_fees = row.get("fees") if isinstance(row.get("fees"), dict) else {}
    fallback_fee = row.get("shippingFee")
    fee_overrides = {
        "margin40": previous_fees.get("margin40", fallback_fee),
        "margin30": previous_fees.get("margin30", previous_fees.get("margin15", fallback_fee)),
        "margin20": previous_fees.get("margin20", fallback_fee),
        "margin10": previous_fees.get("margin10", fallback_fee),
        "breakEven": previous_fees.get("breakEven", fallback_fee),
    }
    results = {
        "margin40": calculate_price(row.get("costRmb"), tier_id, tier, 0.40, config, fee_overrides["margin40"]),
        "margin30": calculate_price(row.get("costRmb"), tier_id, tier, 0.30, config, fee_overrides["margin30"]),
        "margin20": calculate_price(row.get("costRmb"), tier_id, tier, 0.20, config, fee_overrides["margin20"]),
        "margin10": calculate_price(row.get("costRmb"), tier_id, tier, 0.10, config, fee_overrides["margin10"]),
        "breakEven": calculate_price(row.get("costRmb"), tier_id, tier, 0.00, config, fee_overrides["breakEven"]),
    }
    updated = dict(row)
    for key in ("margin40", "margin30", "margin20", "margin10", "breakEven"):
        updated[key] = results[key]["price"]
    for key in ("margin40", "margin30", "margin20", "margin10"):
        updated[f"profit{key.removeprefix('margin')}Rmb"] = results[key]["profit_rmb"]
    updated["shippingFee"] = results["margin30"]["shipping_fee"]
    updated["fees"] = {key: result["shipping_fee"] for key, result in results.items()}
    updated["returnRates"] = {key: result["return_rate"] for key, result in results.items()}
    return updated


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, default=HISTORY_PATH)
    args = parser.parse_args()
    config = read_json(CONFIG_PATH)
    history = read_json(args.source)
    records = history.get("records", [])
    timestamp = time.strftime("%Y%m%d-%H%M%S")
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    backup_path = BACKUP_DIR / f"quick_pricing_history_before_return_rule_{timestamp}.json"
    shutil.copy2(HISTORY_PATH, backup_path)

    sku_count = 0
    zero_return_prices = 0
    standard_return_prices = 0
    for record in records:
        tier_id, tier = tier_from_record(record)
        rows = [recalculate_row(row, tier_id, tier, config) for row in record.get("rows", [])]
        record["rows"] = rows
        record["pricingRuleVersion"] = "returns-above-usd-3-no-disposal-fee"
        sku_count += len(rows)
        for row in rows:
            for rate in row["returnRates"].values():
                if rate == 0:
                    zero_return_prices += 1
                else:
                    standard_return_prices += 1

    history["version"] = max(int(history.get("version", 1)), 2)
    history["pricingRuleVersion"] = "returns-above-usd-3-no-disposal-fee"
    history["recalculatedAt"] = time.time()
    write_json(HISTORY_PATH, history)
    print(
        json.dumps(
            {
                "records": len(records),
                "skus": sku_count,
                "price_points_without_returns": zero_return_prices,
                "price_points_with_returns": standard_return_prices,
                "backup": str(backup_path),
            },
            ensure_ascii=False,
        )
    )


if __name__ == "__main__":
    main()
