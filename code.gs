function exportShopifyOrders() {
  // Configuration
  const CONFIG = {
    shopifyStore: "spalding-bros",
    apiToken: "XXXXXX",
    defaultTaxRate: "4",
    ordersLimit: 250,
    outputFolder: "ORDINI"
  };

  // Setup
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  
  // Define headers
  const headers = [
    "Order Number", "Order Status", "Completed Date", "Customer Note",
    "First Name (Billing)", "Last Name (Billing)", "Company (Billing)", "Address 1&2 (Billing)",
    "City (Billing)", "State Code (Billing)", "Postcode (Billing)", "Country Code (Billing)",
    "Email (Billing)", "Phone (Billing)", "First Name (Shipping)", "Last Name (Shipping)",
    "Address 1&2 (Shipping)", "City (Shipping)", "State Code (Shipping)",
    "Postcode (Shipping)", "Country Code (Shipping)", "Shipping Method Title", "Payment Method Title",
    "Cart Discount Amount", "Order Subtotal Amount", "Order Shipping Amount", 
    "Order Total Tax Amount", "Order Refund Amount", "Order Total Amount", "Item #",
    "SKU", "Item Name", "Quantity", "Item Cost", "Aliquota IVA", "Item Tax Amount", "Item Total Amount", "Coupon Code", "Discount Amount", "Discount Amount Tax",
    "Ragione Sociale", "Partita IVA", "Codice Fiscale", "Codice SDI"
  ];
  sheet.appendRow(headers);

  // Get yesterday's date
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayStr = yesterday.toISOString().split('T')[0];
  
  // Format current date for filename
  const now = new Date();
  const formattedDate = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd-MM-yyyy-HH-mm-ss");
  
  // API request options
  const apiOptions = {
    "method": "GET",
    "headers": {
      "X-Shopify-Access-Token": CONFIG.apiToken,
      "Content-Type": "application/json"
    },
    "muteHttpExceptions": true
  };
  
  // Fetch orders
  const orders = fetchOrders(CONFIG, apiOptions);
  if (!orders || orders.length === 0) {
    Logger.log("Nessun ordine trovato.");
    return;
  }
  
  // Process each order
  orders.forEach(order => {
    // Check if we should include this order
    const { includeOrder, hasRefundToday, refundedItems } = shouldIncludeOrder(order, yesterdayStr, CONFIG, apiOptions);
    
    if (!includeOrder) return;
    
    // Extract order data
    const orderData = extractOrderData(order, CONFIG, apiOptions, yesterdayStr);
    
    // Process line items based on order type
    if (hasRefundToday && !orderData.fulfillmentYesterday) {
      // Process only refunded items
      processRefundedItems(order, refundedItems, orderData, sheet);
    } else {
      // Process all items
      processAllItems(order, refundedItems, orderData, sheet);
    }
  });

  // Save to CSV
  const fileName = saveToCSV(sheet, formattedDate, CONFIG.outputFolder);
  
  // Log riassuntivo finale
  Logger.log(`RIEPILOGO ESPORTAZIONE ORDINI SHOPIFY - ${formattedDate}`);
  Logger.log(`Totale ordini elaborati: ${orders ? orders.length : 0}`);
  Logger.log(`File CSV salvato: ${fileName}`);
}

/**
 * Fetches orders from Shopify API
 */
function fetchOrders(config, apiOptions) {
  try {
    const startDate = new Date();
    startDate.setDate(startDate.getDate() - 10); // Ultimi 5 giorni
    const endDate = new Date(); // Oggi

    const startStr = startDate.toISOString();
    const endStr = endDate.toISOString();

    const apiUrl = `https://${config.shopifyStore}.myshopify.com/admin/api/2025-01/orders.json?status=any&limit=${config.ordersLimit}&created_at_min=${startStr}&created_at_max=${endStr}`;
    
    const response = UrlFetchApp.fetch(apiUrl, apiOptions);
    const data = JSON.parse(response.getContentText());
    return data.orders || [];
  } catch (error) {
    Logger.log(`Error fetching orders: ${error.message}`);
    return [];
  }
}


/**
 * Determines if an order should be included in the export
 */
function shouldIncludeOrder(order, yesterdayStr, config, apiOptions) {
  let includeOrder = false;
  let hasRefundToday = false;
  const refundedItems = {};
  
  // Case 1: Order fulfilled yesterday
  let fulfillmentYesterday = false;
  if (order.fulfillments && order.fulfillments.length > 0) {
    fulfillmentYesterday = order.fulfillments.some(f => {
      const fulfillmentDate = new Date(f.created_at).toISOString().split('T')[0];
      return fulfillmentDate === yesterdayStr;
    });
    
    if (fulfillmentYesterday) {
      includeOrder = true;
    }
  }
  
  // Case 2: Order with refund created yesterday
  try {
    const refundsApiUrl = `https://${config.shopifyStore}.myshopify.com/admin/api/2025-01/orders/${order.id}/refunds.json`;
    const refundsResponse = UrlFetchApp.fetch(refundsApiUrl, apiOptions);
    const refundsData = JSON.parse(refundsResponse.getContentText());
    
    if (refundsData.refunds && refundsData.refunds.length > 0) {
      refundsData.refunds.forEach(refund => {
        const refundDate = new Date(refund.created_at).toISOString().split('T')[0];
        const isRefundedYesterday = refundDate === yesterdayStr;
        
        if (isRefundedYesterday) {
          hasRefundToday = true;
          includeOrder = true;
        }
        
        if (refund.refund_line_items && refund.refund_line_items.length > 0) {
          refund.refund_line_items.forEach(refundItem => {
            if (!refundedItems[refundItem.line_item_id]) {
              refundedItems[refundItem.line_item_id] = {
                quantity: 0, 
                refundedToday: false
              };
            }
            refundedItems[refundItem.line_item_id].quantity += refundItem.quantity;
            
            if (isRefundedYesterday) {
              refundedItems[refundItem.line_item_id].refundedToday = true;
            }
          });
        }
      });
    }
  } catch (error) {
    Logger.log(`Error fetching refunds for order ${order.id}: ${error.message}`);
  }
  
  return { includeOrder, hasRefundToday, refundedItems, fulfillmentYesterday };
}

/**
 * Extracts relevant data from an order
 */
function extractOrderData(order, config, apiOptions, yesterdayStr) {
  const customer = order.customer || {};
  const billing = order.billing_address || {};
  const shipping = order.shipping_address || {};
  
  // Get payment gateway
  let paymentGateway = "N/A";
  try {
    const transactionsApiUrl = `https://${config.shopifyStore}.myshopify.com/admin/api/2025-01/orders/${order.id}/transactions.json`;
    const transactionsResponse = UrlFetchApp.fetch(transactionsApiUrl, apiOptions);
    const transactionsData = JSON.parse(transactionsResponse.getContentText());
    
    if (transactionsData.transactions && transactionsData.transactions.length > 0) {
      const successfulTransaction = transactionsData.transactions.find(t => t.status === "success");
      if (successfulTransaction && successfulTransaction.gateway_display_name) {
        paymentGateway = successfulTransaction.gateway_display_name;
      }
    }
  } catch (error) {
    Logger.log(`Error fetching transactions for order ${order.id}: ${error.message}`);
  }

  // Fallback to payment_gateway_names
  if (paymentGateway === "N/A" && order.payment_gateway_names && order.payment_gateway_names.length > 0) {
    paymentGateway = order.payment_gateway_names.join(", ");
  }

  // Normalize payment gateway name
  if (paymentGateway === "Cash on Delivery (COD)") {
    paymentGateway = "Contrassegno";
  }

  // Extract fiscal data
  const fiscalData = extractFiscalData(order);
  
  // Calculate tax totals
  const { calculatedOrderTaxTotal, subtotalWithoutTax } = calculateOrderTaxes(order);
  
  // Check if order was fulfilled yesterday
  const fulfillmentYesterday = order.fulfillments && order.fulfillments.some(f => {
    const fulfillmentDate = new Date(f.created_at).toISOString().split('T')[0];
    return fulfillmentDate === yesterdayStr;
  });
  
  return {
    customer,
    billing,
    shipping,
    paymentGateway,
    orderStatus: order.financial_status,
    fiscalData,
    calculatedOrderTaxTotal,
    subtotalWithoutTax,
    fulfillmentYesterday
  };
}

/**
 * Extracts fiscal data from order attributes
 */
function extractFiscalData(order) {
  let ragioneSociale = "";
  let partitaIVA = "";
  let codiceFiscale = "";
  let codiceSDI = "";
  
  if (order.note_attributes && order.note_attributes.length > 0) {
    order.note_attributes.forEach(attr => {
      if (attr.name === "Ragione Sociale") ragioneSociale = attr.value || "";
      if (attr.name === "Partita IVA") partitaIVA = attr.value || "";
      if (attr.name === "Codice Fiscale") codiceFiscale = attr.value || "";
      if (attr.name === "Codice SDI") codiceSDI = attr.value || "";
    });
  }
  
  return { ragioneSociale, partitaIVA, codiceFiscale, codiceSDI };
}

/**
 * Calculates tax totals for an order
 */
function calculateOrderTaxes(order) {
  // Verifica se l'ordine ha già i totali fiscali da Shopify
  let shopifyTaxTotal = 0;
  let shopifySubtotalWithoutTax = 0;
  let usingShopifyTax = false;
  
  // Cerca dati fiscali diretti nell'ordine
  if (order.tax_lines && order.tax_lines.length > 0) {
    order.tax_lines.forEach(taxLine => {
      if (taxLine.price) {
        shopifyTaxTotal += parseFloat(taxLine.price);
      }
    });
    usingShopifyTax = true;
  }
  
  // Cerca i subtotali netti nell'ordine
  if (order.current_subtotal_price_set && 
      order.current_subtotal_price_set.presentment_money && 
      order.current_subtotal_price && 
      order.current_total_tax) {
    shopifySubtotalWithoutTax = parseFloat(order.current_subtotal_price) - parseFloat(order.current_total_tax);
    usingShopifyTax = true;
  } else if (order.subtotal_price) {
    shopifySubtotalWithoutTax = parseFloat(order.subtotal_price);
    usingShopifyTax = true;
  }
  
  // Se abbiamo i dati fiscali da Shopify, utilizzali direttamente
  if (usingShopifyTax && shopifyTaxTotal > 0) {
    return {
      calculatedOrderTaxTotal: shopifyTaxTotal,
      subtotalWithoutTax: shopifySubtotalWithoutTax
    };
  }
  
  // Fallback al nostro calcolo se i dati fiscali non sono disponibili
  let calculatedOrderTaxTotal = 0;
  let subtotalWithoutTax = 0;
  
  // Raggruppa gli articoli per aliquota IVA
  const taxGroups = {};
  const excludedTitles = ["Commissione pagamento alla consegna"];
  
  // Prima fase: raggruppamento degli articoli per aliquota IVA
  order.line_items.forEach(item => {
    if (excludedTitles.some(title => item.title.includes(title))) return;
    
    const taxRate = getTaxRate(item);
    const itemPriceWithTax = parseFloat(item.price) || 0;
    const quantity = item.quantity || 0;
    const totalPriceWithTax = itemPriceWithTax * quantity;
    
    if (!taxGroups[taxRate]) {
      taxGroups[taxRate] = 0;
    }
    
    taxGroups[taxRate] += totalPriceWithTax;
  });
  
  // Seconda fase: calcolo dell'imponibile e dell'IVA per ciascun gruppo
  for (let taxRate in taxGroups) {
    const totalWithTax = taxGroups[taxRate];
    
    // Formula corretta: Totale * 100 / (100 + aliquota)
    const netAmount = totalWithTax * 100 / (100 + parseFloat(taxRate));
    const taxAmount = totalWithTax - netAmount;
    
    subtotalWithoutTax += netAmount;
    calculatedOrderTaxTotal += taxAmount;
  }
  
  return {
    calculatedOrderTaxTotal,
    subtotalWithoutTax
  };
}

/**
 * Processes refunded items for an order
 */
function processRefundedItems(order, refundedItems, orderData, sheet) {
  order.line_items.forEach((item, index) => {
    if (item.title.includes("Commissione pagamento alla consegna")) return;
    
    if (refundedItems[item.id] && refundedItems[item.id].refundedToday) {
      const itemData = calculateItemData(item, refundedItems[item.id].quantity);
      appendRowToSheet(sheet, order, item, itemData, orderData, index, -refundedItems[item.id].quantity);
    }
  });
}

/**
 * Processes all items for an order
 */
function processAllItems(order, refundedItems, orderData, sheet) {
  order.line_items.forEach((item, index) => {
    if (item.title.includes("Commissione pagamento alla consegna")) return;
    
    if (refundedItems[item.id]) {
      // Process remaining items
      const refundedQuantity = refundedItems[item.id].quantity;
      const remainingQuantity = item.quantity - refundedQuantity;
      
      if (remainingQuantity > 0) {
        const itemData = calculateItemData(item, remainingQuantity);
        appendRowToSheet(sheet, order, item, itemData, orderData, index, remainingQuantity);
      }
      
      // Process refunded items
      const refundedItemData = calculateItemData(item, refundedQuantity);
      appendRowToSheet(sheet, order, item, refundedItemData, orderData, index, -refundedQuantity);
    } else {
      // Process normal items
      const itemData = calculateItemData(item, item.quantity);
      appendRowToSheet(sheet, order, item, itemData, orderData, index, item.quantity);
    }
  });
}

/**
 * Calculates data for an item
 */
function calculateItemData(item, quantity) {
  // Se l'articolo ha già i dati fiscali calcolati da Shopify, utilizzarli direttamente
  if (item.tax_lines && item.tax_lines.length > 0) {
    const taxLine = item.tax_lines[0];
    const taxRate = getTaxRate(item);
    const itemPriceWithTax = parseFloat(item.price) || 0;
    const absQuantity = Math.abs(quantity);
    
    // Ottieni l'importo dell'IVA direttamente da Shopify
    let taxAmount = 0;
    if (taxLine.price) {
      taxAmount = parseFloat(taxLine.price);
    } else if (taxLine.rate && item.pre_tax_price) {
      // Se abbiamo il tasso e il prezzo pre-tassa
      taxAmount = parseFloat(item.pre_tax_price) * parseFloat(taxLine.rate);
    }
    
    // Calcola il prezzo netto unitario (prezzo con IVA - importo IVA unitario)
    const itemTaxAmount = taxAmount / absQuantity;
    const itemNetPrice = itemPriceWithTax - itemTaxAmount;
    
    return {
      taxRate,
      taxAmount: taxAmount,
      netPrice: itemNetPrice,
      totalAmount: itemPriceWithTax * absQuantity
    };
  }
  
  // Fallback alla nostra formula se Shopify non fornisce i dati fiscali
  const taxRate = getTaxRate(item);
  const itemPriceWithTax = parseFloat(item.price) || 0;
  const absQuantity = Math.abs(quantity);
  const totalPriceWithTax = itemPriceWithTax * absQuantity;
  
  // Calcola l'imponibile totale: prezzo lordo * 100 / (100 + aliquota IVA)
  const totalNetAmount = totalPriceWithTax * 100 / (100 + parseFloat(taxRate));
  const totalTaxAmount = totalPriceWithTax - totalNetAmount;
  const itemNetPrice = totalNetAmount / absQuantity;
  
  return {
    taxRate,
    taxAmount: totalTaxAmount,
    netPrice: itemNetPrice,
    totalAmount: totalPriceWithTax
  };
}

/**
 * Appends a row to the sheet
 */
function appendRowToSheet(sheet, order, item, itemData, orderData, index, quantity) {
  const { customer, billing, shipping, paymentGateway, orderStatus, fiscalData, calculatedOrderTaxTotal, subtotalWithoutTax } = orderData;
  const { ragioneSociale, partitaIVA, codiceFiscale, codiceSDI } = fiscalData;
  
  // Utilizza direttamente i dati calcolati da calculateItemData
  const itemNetPrice = itemData.netPrice;
  const itemTaxAmount = itemData.taxAmount / Math.abs(quantity);
  const itemTotalPrice = parseFloat(item.price) || 0;
  
  sheet.appendRow([
    order.name,
    orderStatus,
    (order.fulfillments && order.fulfillments.length > 0)
    ? order.fulfillments[0].created_at
    : order.created_at,
    customer.note || "",
    billing.first_name || "", billing.last_name || "", billing.company || "", (billing.address1 || "") + " " + (billing.address2 || ""),
    billing.city || "", billing.province_code || "", billing.zip || "", billing.country_code || "",
    customer.email || "", (shipping.phone || "").replace(/[^\d]/g, ""),
    shipping.first_name || "", shipping.last_name || "", (shipping.address1 || "") + " " + (shipping.address2 || ""),
    shipping.city || "", shipping.province_code || "", shipping.zip || "", shipping.country_code || "",
    (order.shipping_lines && order.shipping_lines.length > 0) ? order.shipping_lines.map(s => s.title).join(", ") : "",
    paymentGateway,
    formatMonetaryValue(order.discount_codes && order.discount_codes.length > 0 ? order.discount_codes[0].amount : "0"),
    formatMonetaryValue(subtotalWithoutTax),
    formatMonetaryValue(order.total_shipping_price_set && order.total_shipping_price_set.presentment_money ? order.total_shipping_price_set.presentment_money.amount : "0"),
    formatMonetaryValue(calculatedOrderTaxTotal),
    formatMonetaryValue(order.total_refunds || "0"),
    formatMonetaryValue(order.total_price || "0"),
    index + 1,
    item.sku || "",
    item.title,
    quantity,
    formatMonetaryValue(itemNetPrice),
    itemData.taxRate,
    formatMonetaryValue(itemTaxAmount),
    formatMonetaryValue(itemTotalPrice * Math.sign(quantity)),
    (order.discount_codes && order.discount_codes.length > 0) ? order.discount_codes[0].code : "",
    formatMonetaryValue(order.discount_codes && order.discount_codes.length > 0 ? order.discount_codes[0].amount : "0"),
    "0,00",
    ragioneSociale, partitaIVA, codiceFiscale, codiceSDI
  ]);
}

/**
 * Formats a monetary value to always have 2 decimal places with comma as decimal separator
 */
function formatMonetaryValue(value) {
  // Gestione valori nulli o undefined
  if (value === null || value === undefined) {
    return "0,00";
  }
  
  // Conversione a numero se è una stringa
  let numValue = typeof value === 'string' ? parseFloat(value.replace(',', '.')) : value;
  
  // Controllo se è un numero valido
  if (isNaN(numValue)) {
    return "0,00";
  }
  
  // Arrotondamento a 2 decimali per evitare errori di precisione
  numValue = Math.round(numValue * 100) / 100;
  
  // Formattazione con 2 decimali e sostituzione del punto con la virgola
  return numValue.toFixed(2).replace('.', ',');
}

/**
 * Saves the sheet to a CSV file
 */
function saveToCSV(sheet, formattedDate, folderName) {
  try {
    const folder = DriveApp.getFoldersByName(folderName).next();
    const fileName = "ORDINIWEB.csv";
    const csvContent = convertSheetToCSV(sheet);
    // Converti in Windows-1252
    const blob = Utilities.newBlob(csvContent, "text/csv", fileName).getAs('text/csv');
    blob.setContentTypeFromExtension();
    const file = folder.createFile(blob);
    Logger.log("File CSV creato: " + file.getUrl());
    return file.getUrl();
  } catch (error) {
    Logger.log(`Error saving CSV: ${error.message}`);
    return null;
  }
}

/**
 * Gets the tax rate for an item
 */
function getTaxRate(item) {
  let taxRate = "";
  
  // Method 1: Get rate from tax_lines
  if (item.tax_lines && item.tax_lines.length > 0) {
    const taxLine = item.tax_lines[0];
    if (taxLine.rate) {
      taxRate = (taxLine.rate * 100).toFixed(0);
    } else if (taxLine.title) {
      const match = taxLine.title.match(/\d+/);
      if (match) {
        taxRate = match[0];
      }
    }
  }
  
  // Method 2: Calculate from price and tax amount
  if (!taxRate && item.price && item.tax_amount && parseFloat(item.price) > 0) {
    const taxPercentage = (parseFloat(item.tax_amount) / parseFloat(item.price)) * 100;
    taxRate = taxPercentage.toFixed(0);
  }
  
  // Method 3: Look in item properties
  if (!taxRate && item.properties) {
    const taxProperty = item.properties.find(prop => 
      prop.name && (
        prop.name.toLowerCase().includes("tax") || 
        prop.name.toLowerCase().includes("iva") || 
        prop.name.toLowerCase().includes("aliquota")
      )
    );
    
    if (taxProperty) {
      const match = taxProperty.value.match(/\d+/);
      if (match) {
        taxRate = match[0];
      } else {
        taxRate = taxProperty.value;
      }
    }
  }
  
  // Default to 4% if no rate found
  if (!taxRate || taxRate === "") {
    taxRate = "4";
  }
  
  return taxRate;
}

/**
 * Converts a sheet to CSV format
 */
function convertSheetToCSV(sheet) {
  const data = sheet.getDataRange().getValues();
  // Identificare gli indici delle colonne che contengono valori monetari
  const monetaryColumnIndices = [
    23, // Cart Discount Amount
    24, // Order Subtotal Amount
    25, // Order Shipping Amount
    26, // Order Total Tax Amount
    27, // Order Refund Amount
    28, // Order Total Amount
    33, // Item Cost
    35, // Item Tax Amount
    36, // Item Total Amount
    38, // Discount Amount
    39  // Discount Amount Tax
  ];
  
  let csvContent = "";
  data.forEach((row, rowIndex) => {
    // Join with semicolon and ensure each cell is properly formatted
    if (rowIndex === 0) return; // Salta l'intestazione
    csvContent += row.map((cell, colIndex) => {
      
      // Per i valori monetari, assicuriamo la formattazione corretta
      if (monetaryColumnIndices.includes(colIndex)) {
        // Conversione a stringa se necessario
        let value = typeof cell === 'string' ? cell : cell.toString();
        
        // Assicuriamo che sia nel formato con virgola
        if (value.includes('.')) {
          value = value.replace('.', ',');
        }
        
        // Assicuriamo che ci siano sempre 2 decimali
        if (!value.includes(',')) {
          value += ',00';
        } else {
          const parts = value.split(',');
          if (parts[1].length === 1) {
            value += '0';
          }
        }
        
        return value;
      }
      
      // Per i valori postali e numerici non monetari, assicurati che non ci siano decimali
      // ZIP/CAP (indici 10 e 19)
      if (typeof cell === 'number' && (colIndex === 10 || colIndex === 19)) {
        return Math.round(cell).toString();
      }
      
      return cell;
    }).join(";") + "\r\n"; // MS-DOS line ending
  });
  return csvContent;
}
