WITH 
suppliers_clean AS (
    SELECT id, name, code
    FROM `wild-earth`.suppliers
    WHERE obsolete = 0
),
base_stock AS (
    SELECT *
    FROM `wild-earth`.stocks
    WHERE warehouse_code = 'W1' AND min_alert > 0
),
valid_products AS (
    SELECT *
    FROM `wild-earth`.products
    WHERE obsolete = 0
      AND supplier_id NOT IN (
          '19', 
          '134', 
          '267', 
          '5', 
          '209',
          '531'
      )
	AND brand <> 'Lowe Alpine'
	AND NOT (brand = 'Rab' AND supplier_id = 90)
	AND (
		product_group IS NULL 
		OR product_group NOT IN ('Gift Vouchers', 'GWP', 'Miscellaneous', 'Admin')
      )
),
purchase_orders_placed AS (
    SELECT *
    FROM `wild-earth`.purchase_orders
    WHERE warehouse_code = 'W1' 
      AND order_status IN ('Placed', 'Costed')
),
po_lines AS (
    SELECT 
        pol.product_id,
        SUM(pol.order_quantity) AS total_order_quantity
    FROM `wild-earth`.purchase_order_lines pol
    JOIN purchase_orders_placed po 
        ON pol.purchase_order_id = po.id
    GROUP BY pol.product_id
),
final_data AS (
    SELECT  
        p.product_code AS `SKU`,
        p.description AS `Name`,
        p.qty_on_hand AS 'TotalSOH',
        sc.name AS `Supplier`,
        sc.id AS `Sup.Id`,
        p.default_purchase_price AS `Cost`,
        p.barcode AS `Barcode`,
        COALESCE(s_w1.max_alert, 0) AS `MaxW1`,
        COALESCE(s_w1.qty_on_hand, 0) AS `w1_SOH`,
        COALESCE(pol.total_order_quantity, 0) AS `OnPurchase`,
        COALESCE(s_w1.max_alert, 0) - (COALESCE(pol.total_order_quantity, 0) + COALESCE(s_w1.qty_on_hand, 0)) AS `QtyOrdered`,
        p.pack_size AS `PackSize`
    FROM base_stock s_w1
    INNER JOIN valid_products p 
        ON s_w1.product_id = p.id
    LEFT JOIN po_lines pol 
        ON p.id = pol.product_id
    LEFT JOIN suppliers_clean sc 
        ON p.supplier_id = sc.id
	WHERE (COALESCE(pol.total_order_quantity, 0) + COALESCE(s_w1.qty_on_hand, 0)) < COALESCE(s_w1.min_alert, 0)
    AND (sc.name NOT LIKE '%INT12-24%')
	AND (sc.name NOT LIKE '%(N.-R)%')
	AND (sc.name NOT LIKE '%(F.-D)%')
    AND (sc.name NOT LIKE '%(V.-T)%')
)
SELECT *
FROM final_data
ORDER BY Supplier ASC
LIMIT 40000;