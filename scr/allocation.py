from __future__ import annotations

from decimal import Decimal
from typing import List

from scr.text_utils import quantize_2


def allocate_proportional(total_amount: Decimal, shares: List[Decimal]) -> List[Decimal]:
    """按份额比例分配，保留两位小数。

    - 最后一个分配项接收尾差，保证合计等于总额。
    - 份额合计为 0 或份额为空时，返回 0 列表。
    """
    if total_amount is None:
        total_amount = Decimal("0")

    if not shares:
        return []

    total_shares = sum((s or Decimal("0")) for s in shares)
    if total_shares == 0:
        return [Decimal("0.00") for _ in shares]

    allocated: List[Decimal] = []
    running_sum = Decimal("0.00")
    for idx, s in enumerate(shares):
        if idx == len(shares) - 1:
            # 最后一个接收尾差
            last = total_amount - running_sum
            allocated.append(quantize_2(last))
        else:
            part = (total_amount * (s or Decimal("0"))) / total_shares
            q = quantize_2(part)
            allocated.append(q)
            running_sum += q
    return allocated
