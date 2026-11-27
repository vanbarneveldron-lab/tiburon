# Velles i Novelles Apps Script (DEV)

## ItemTypes sheet schema

The current DEV ItemTypes tab uses these columns in this order:

| Column | Purpose |
| --- | --- |
| item_type_id | Stable key stored on items; do not rename or change existing IDs. |
| name_ca | Catalan display name. |
| active | Boolean flag (TRUE to show the type, FALSE to hide it). |
| name_es | Spanish display name. |
| name_en | English display name. |

The new-item and inventory-filter dropdowns display the name for the configured APP_LANGUAGE (ca, es, or en). If that translation is blank, the app falls back to name_ca, then to item_type_id. When adding an item type, keep its ID stable and fill in all three names. Existing item IDs and active flags should be preserved during sheet migrations.


## Sales module (DEV)

A sale is created from an active inventory item. Its channel choices are the active platforms on which that specific item has an active publication, plus an other/in-store option. Each sale stores the item ID and reference, ISO sale date, sales channel, VAT-inclusive gross price, VAT rate, VAT amount, net amount, invoice placeholders, and status in the existing `Sales` sheet schema. The VAT rate defaults to 4% in the form and accepts a rate from 0% to 100% with up to two decimal places. When an item is sold, all its active publications move to `remove_needed`. Each can then be manually confirmed as withdrawn; the app records the confirmation date in `withdrawn_at` and keeps the publication record for history. Sales can be reversed: the sale row remains for the audit trail with status `reversed`, the item returns to `active`, and pending publication removals return to `active`. Reversed sales are retained in history and excluded from the active sales count.

The Publications screen records manual listings against active inventory items and configured active Platforms, with listing and withdrawal dates. Withdrawal keeps the record for history. The Reports screen summarizes active sales between optional inclusive start and end dates (defaulting to the current year through today), including gross, VAT, net, and count. Reversed sales are excluded.
