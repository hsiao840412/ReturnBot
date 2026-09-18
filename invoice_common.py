"""Invoice placement shared by normal KBB and consignment recall."""
from dataclasses import dataclass


@dataclass(frozen=True)
class InvoiceLayout:
    count: int
    start: int = 13
    default_rows: int = 3

    @property
    def delta(self):
        return self.count - self.default_rows

    @property
    def last(self):
        return self.start + self.count - 1

    @property
    def total_row(self):
        return 16 + self.delta

    @property
    def quantity_row(self):
        return 18 + self.delta


def invoice_detail(index, part, rma, description, quantity, returns, unit_price, total):
    # D:G is the description area in the KBB template. Carton assignment is
    # completed during packing, as in the existing normal KBB workflow.
    return [index, part, rma, description, None, None, None,
            quantity, returns, unit_price, total, None]
