#!/usr/bin/env python3
"""Apply verified Seller Central fulfillment fees to September refurb SKUs."""

from __future__ import annotations

import json
import shutil
import time
from pathlib import Path

from recalculate_quick_history import (
    calculate_price,
    read_json,
    tier_from_record,
    write_json,
)


ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "pricing_config.json"
HISTORY_PATH = ROOT / "config" / "quick_pricing_history.json"
BACKUP_DIR = ROOT / "config" / "backups"
RECORD_ID = "return-refurb-new-products-20260928-under4"

# Seller Central > Manage All Inventory > Amazon fulfillment fee, pages 1-3.
BACKEND_FEES = {
    "TTCA-PhonePouch-Black": 1.77,
    "TTCA-PhonePouch-White": 1.77,
    "TTCA-CyclingShorts-2XL": 0.99,
    "TTCA-CyclingShorts-M": 0.99,
    "TTCA-CrystalBracelet-Pearl": 0.50,
    "TTCA-SpiceShaker-Blue-2pcs": 2.60,
    "TTCA-CleaningBrush-WhiteGrey-4pcs": 0.50,
    "TTCA-CherryCharm-Red": 0.50,
    "TTCA-ScarfRing-3pcs": 1.77,
    "TTCA-ResinKeychain-Wood": 0.50,
    "TTCA-IncenseHolder-1pc": 0.50,
    "TTCA-CapybaraKeychain-3pcs": 0.50,
    "TTCA-DogToy-RopeKnot-1pc": 1.77,
    "TTCA-DogToy-MixedSet-3pcs": 2.60,
    "TTCA-DogToy-BallSet-2pcs": 1.77,
    "TTCA-BakingPaper-50pcs": 1.77,
    "TTCA-MarkingRuler-Red": 0.88,
    "TTCA-LaundryBag-White": 2.60,
    "TTCA-ShoeBag-Yellow": 3.22,
    "TTCA-LeatherPunch-1pc": 2.60,
    "TTCA-ExerciseBand-1pc": 2.60,
    "TTCA-BottleOpener-3pcs": 0.50,
    "TTCA-RugGripper-16pcs": 0.88,
    "TTCA-SnapHook-20pcs-40mm": 0.50,
    "TTCA-MagneticMount-1pc": 0.50,
    "TTCA-SewingKit-27pcs": 0.50,
    "TTCA-HexKey-9pcs": 0.88,
    "TTCA-TweezerSet-4pcs": 0.50,
    "TTCA-HairTie-DoubleArc-Silver": 1.77,
    "TTCA-HairTie-Bow-Gold": 0.88,
    "TTCA-HairTie-Bow-Silver": 0.88,
    "TTCA-HairTie-DoubleArc-Gold": 0.88,
    "TTCA-MeshStrainer-200mesh": 0.50,
    "TTCA-SportSocks-5pairs": 2.84,
    "TTCA-RepairKit-23pcs": 0.88,
    "TTCA-BoatSocks-7pairs": 3.48,
    "TTCA-MagnetizerTool-Green": 0.88,
    "TTCA-AnkleSocks-5pairs": 2.84,
    "TTCA-CottonUnderwear-M-5pcs": 2.84,
    "TTCA-NeckGaiter-LightSet": 0.99,
    "TTCA-NeckGaiter-DarkSet": 0.99,
    "TTCA-ButterflyUnderwear-M": 2.05,
    "TTCA-ButterflyUnderwear-S": 0.99,
    "TTCA-CamoMask-BlueCamo": 2.05,
    "TTCA-CamoMask-GreenCamo": 1.77,
    "TTCA-ToiletCover-1pc": 0.88,
    "TTCA-GradientSleeve-BlueYellow": 0.50,
    "TTCA-GradientSleeve-BluePink": 0.50,
    "TTCA-GradientSleeve-PinkYellow": 0.50,
    "TTCA-CoolingSleeve-Glacier": 0.50,
    "TTCA-IceTray-1pc": 1.77,
    "TTCA-BentoBox-2pcs": 3.72,
    "TTCA-SwabHolder-Clear": 1.77,
    "TTCA-ScrewdriverSet-Green-22in1": 1.77,
    "TTCA-EggMold-Silver-4pcs": 1.77,
    "TTCA-CleaningSponge-Blue": 1.77,
    "TTCA-JarSpatula-Multicolor-4pcs": 0.88,
    "TTCA-CleaningScraper-Green-5blades": 1.77,
    "TTCA-MakeupSponge-Pink-14pcs": 1.77,
    "TTCA-StorageRack-Black": 4.42,
    "TTCA-ResistanceBand-Multicolor-5pcs": 0.88,
    "TTCA-KitchenSpatula-Black": 1.77,
    "TTCA-HairClip-Pink": 0.50,
    "TTCA-HairClip-Beige": 0.50,
    "TTCA-HairClip-Brown": 0.50,
    "TTCA-HairClip-Black": 0.50,
    "TTCA-SpatulaTongs-Silver": 0.50,
    "TTCA-YogaUnderwear-XL": 3.48,
    "TTCA-YogaUnderwear-L": 2.05,
    "TTCA-LaceUnderwear-M": 2.05,
    "TTCA-LaceUnderwear-S": 2.05,
    "TTCA-IceSilk-L": 0.99,
    "TTCA-FaceMask-BrightSet": 2.05,
    "TTCA-FaceMask-DarkSet": 0.99,
    "TTCA-BakingMat-Black-2pcs": 1.77,
    "TTCA-OilSeparator-Silver": 1.77,
    "TTCA-PhoneStand-Gray": 1.77,
    "TTCA-PhoneStand-Silver": 0.50,
}

PRICE_KEYS = ("margin40", "margin30", "margin20", "margin10", "breakEven")
MARGINS = {
    "margin40": 0.40,
    "margin30": 0.30,
    "margin20": 0.20,
    "margin10": 0.10,
    "breakEven": 0.00,
}


def recalculate_with_backend_fee(row, tier_id, tier, config, backend_fee):
    results = {
        key: calculate_price(
            row.get("costRmb"),
            tier_id,
            tier,
            margin,
            config,
            shipping_fee_override=backend_fee,
        )
        for key, margin in MARGINS.items()
    }
    updated = dict(row)
    for key in PRICE_KEYS:
        updated[key] = results[key]["price"]
    for key in ("margin40", "margin30", "margin20", "margin10"):
        suffix = key.removeprefix("margin")
        updated[f"profit{suffix}Rmb"] = results[key]["profit_rmb"]
    updated["shippingFee"] = backend_fee
    updated["fees"] = {key: backend_fee for key in PRICE_KEYS}
    updated["returnRates"] = {key: results[key]["return_rate"] for key in PRICE_KEYS}
    updated["shippingFeeSource"] = "seller-central-backend"
    return updated


def main():
    config = read_json(CONFIG_PATH)
    history = read_json(HISTORY_PATH)
    record = next((item for item in history["records"] if item.get("id") == RECORD_ID), None)
    if record is None:
        raise RuntimeError(f"未找到记录：{RECORD_ID}")

    timestamp = time.strftime("%Y%m%d-%H%M%S")
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    backup_path = BACKUP_DIR / f"quick_pricing_history_before_backend_shipping_{timestamp}.json"
    shutil.copy2(HISTORY_PATH, backup_path)

    tier_id, tier = tier_from_record(record)
    found_skus = set()
    changed_fee_skus = []
    unchanged_fee_skus = []
    changed_price_points = 0
    updated_rows = []

    for row in record.get("rows", []):
        sku = row.get("sku")
        if sku not in BACKEND_FEES:
            updated_rows.append(row)
            continue

        found_skus.add(sku)
        backend_fee = float(BACKEND_FEES[sku])
        old_fee = float(row.get("shippingFee") or 0)
        if abs(old_fee - backend_fee) > 1e-9:
            changed_fee_skus.append(sku)
        else:
            unchanged_fee_skus.append(sku)

        updated = recalculate_with_backend_fee(row, tier_id, tier, config, backend_fee)
        changed_price_points += sum(
            1
            for key in PRICE_KEYS
            if abs(float(row.get(key) or 0) - float(updated.get(key) or 0)) > 1e-9
        )
        updated_rows.append(updated)

    record["rows"] = updated_rows
    assumptions = record.setdefault("assumptions", {})
    for key in (
        "currency_rate_rmb_to_usd",
        "referral_fee_rate",
        "return_rate",
        "return_rate_price_threshold_usd",
        "first_leg_shipping_usd",
        "target_margin_min",
        "target_price_gap_usd",
        "dimensional_weight_divisor",
        "disposal_fee_usd",
    ):
        if key in config:
            assumptions[key] = config[key]
    record["backendShippingFeeAudit"] = {
        "source": "Amazon Seller Central Manage All Inventory / Amazon fulfillment fee",
        "verifiedPages": [1, 2, 3],
        "verifiedAt": time.time(),
        "verifiedSkuCount": len(found_skus),
    }

    history["recalculatedAt"] = time.time()
    write_json(HISTORY_PATH, history)

    missing_in_record = sorted(set(BACKEND_FEES) - found_skus)
    print(
        json.dumps(
            {
                "record": record.get("name"),
                "record_skus": len(record.get("rows", [])),
                "backend_skus": len(BACKEND_FEES),
                "verified_skus_updated": len(found_skus),
                "shipping_fees_changed": len(changed_fee_skus),
                "shipping_fees_confirmed_unchanged": len(unchanged_fee_skus),
                "price_points_changed": changed_price_points,
                "record_skus_not_yet_verified": len(record.get("rows", [])) - len(found_skus),
                "backend_skus_missing_in_record": missing_in_record,
                "backup": str(backup_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )


if __name__ == "__main__":
    main()
