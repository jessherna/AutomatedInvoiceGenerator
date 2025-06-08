import shutil
from pathlib import Path
import pytest
from invoice import LoadOrders

def test_load_orders(tmp_path):
    """
    Test that the LoadOrders function can load the orders from the sample file.
    """
    sample_src = Path("tests/data/orders_sample_with_id.xlsx")
    sample = tmp_path / "orders_sample_with_id.xlsx"
    shutil.copy(sample_src, sample)

    # Act
    table = LoadOrders(str(sample))

    # Assert type & length
    assert isinstance(table, list)
    assert len(table) == 10

    # Assert first row’s fields
    first = table[0]
    assert first["ItemID"] == "ABC123"
    assert first["Qty"] == 2
    assert first["Price"] == 9.99
