#!/usr/bin/env python3
from pathlib import Path

import generate_pricing
import validate_inputs


ROOT = Path(__file__).resolve().parents[1]


def main():
    validate_inputs.main()
    generate_pricing.main()

    print("")
    print("定价引擎已生成结果：")
    print(f"- 校验报告：{ROOT / 'output' / 'input_validation_report.md'}")
    print(f"- 定价结果：请查看 {ROOT / 'output'} 里最新的 *_自动定价结果.xlsx")
    print(f"- 定价报告：请查看 {ROOT / 'output'} 里最新的 *_自动定价报告.md")


if __name__ == "__main__":
    main()
