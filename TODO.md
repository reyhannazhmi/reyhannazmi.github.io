# TODO: Fix automatic stock madu reduction on penjualan transaksi

## Plan Steps:
- [ ] 1. Update `/input_transaksi` POST handler in sia.py for `jenis_transaksi == 'Penjualan'`:
  - Parse `product_{i}` (item_code), `quantity_{i}` from form (i=1,2,... until missing).
  - For each: lookup item in inventory_data by item_code; validate exists & qty <= stock.
  - Calc total_selling = selling_price * qty, total_hpp = cost_price * qty.
  - Auto-append 4 journal rows: Dr '1-1100 - Kas' total_selling; Cr '4-4000 - Penjualan barang dagang' total_selling; Dr '5-5000 - Harga pokok penjualan' total_hpp; Cr '1-1300 - Persediaan barang dagang' total_hpp.
  - keterangan += f" Penjualan {item.name} {qty} unit".
  - update_inventory_stock(item.name, -qty).
  - Skip manual debit/kredit processing for Penjualan (or validate balance later).
- [ ] 2. For `jenis_transaksi == 'Pembelian'`:
  - Parse `purchase_product_{i}`, `purchase_quantity_{i}`.
  - total_cost = cost_price * qty.
  - Auto-append: Dr '1-1300 - Persediaan barang dagang' total_cost; Cr '2-2100 - Hutang dagang' total_cost.
  - keterangan += f" Pembelian {item.name} {qty} unit".
  - update_inventory_stock(item.name, +qty).
- [ ] 3. For 'Lainnya': keep existing manual debit/kredit + old stock heuristic logic.
- [ ] 4. After all updates, reload inventory_data=load_inventory() for template display.
- [ ] 5. Add validation: total_debit == total_kredit after auto-entries; error if manual used with jenis_transaksi.
- [ ] 6. Test: run app, input penjualan, check inventory.xlsx stock decreased, journal.xlsx entries correct.

Next step: Implement step 1-4 via edit_file on sia.py.
