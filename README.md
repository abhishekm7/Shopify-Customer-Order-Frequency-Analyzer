// Shopify <> Google Sheet Connector Developed by Abhishek Maity.
// For more such micro tool contact abhishek.maity@adbuffs.com
// Linkedin - https://www.linkedin.com/in/abhishek-maity-adbuffs/
// How to find accessToken = Go to Shopify > Apps > Develop Apps > Assign all 'read_' access > Copy Admin API access token
// How to find shopName = Go to Shopify Home > Check Browser URL > Copy the part after 'https://admin.shopify.com/store/'
// If store url is https://admin.shopify.com/store/mydemostore then shopName = mydemostore

function getCustomerOrderFrequency() {
  var accessToken = 'Shopify Access Token';
  var shopName = 'ShopName';
  var X = 365;  // Number of days; you can change this value as needed.

  var date = new Date();
  date.setDate(date.getDate() - X);
  var createdAtMin = date.toISOString();

  var baseUrl = 'https://' + shopName + '.myshopify.com/admin/api/2024-01/orders.json?status=any&created_at_min=' + createdAtMin + '&limit=250';
  var orders = [];
  var hasMoreOrders = true;
  var url = baseUrl;

  while (hasMoreOrders) {
    var options = {
      'method': 'GET',
      'contentType': 'application/json',
      'headers': {
        'X-Shopify-Access-Token': accessToken
      }
    };

    var response = UrlFetchApp.fetch(url, options);
    var data = JSON.parse(response.getContentText());
    orders = orders.concat(data.orders);

    var linkHeader = response.getHeaders()['Link'];
    if (linkHeader && linkHeader.includes('rel="next"')) {
      var match = linkHeader.match(/<([^>]+)>;\s*rel="next"/);
      if (match) {
        url = match[1];
      } else {
        hasMoreOrders = false;
      }
    } else {
      hasMoreOrders = false;
    }
  }

  var customerData = {};

  orders.forEach(function(order) {
    var customerId = order.customer ? order.customer.id : 'N/A';
    var orderSubtotal = parseFloat(order.subtotal_price);

    if (!customerData[customerId]) {
      customerData[customerId] = {
        customerId: customerId,
        orders: []
      };
    }
    customerData[customerId].orders.push({
      orderId: order.id,
      createdAt: new Date(order.created_at),
      orderSubtotal: orderSubtotal,
      products: order.line_items.map(function(item) { return item.name; }),
    });
  });

  var trendSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CustomerOrderFrequency');
  if (!trendSheet) {
    trendSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('CustomerOrderFrequency');
  }
  trendSheet.clear();

  var headers = [
    'Customer ID', 'Number of Orders', 'First Order Date', 'Second Order Date', 'Third Order Date', 'Fourth Order Date',
    'Average Time 1st to 2nd Order', 'Average Time 2nd to 3rd Order', 'Average Time 3rd to 4th Order'
  ];
  trendSheet.appendRow(headers);

  for (var id in customerData) {
    var customer = customerData[id];
    customer.orders.sort(function(a, b) { return a.createdAt - b.createdAt; });

    var firstOrderDate = customer.orders[0] ? customer.orders[0].createdAt : 'N/A';
    var secondOrderDate = customer.orders[1] ? customer.orders[1].createdAt : 'N/A';
    var thirdOrderDate = customer.orders[2] ? customer.orders[2].createdAt : 'N/A';
    var fourthOrderDate = customer.orders[3] ? customer.orders[3].createdAt : 'N/A';

    var avgTime1to2 = customer.orders[1] ? (customer.orders[1].createdAt - customer.orders[0].createdAt) / (1000 * 60 * 60 * 24) : 'N/A';
    var avgTime2to3 = customer.orders[2] ? (customer.orders[2].createdAt - customer.orders[1].createdAt) / (1000 * 60 * 60 * 24) : 'N/A';
    var avgTime3to4 = customer.orders[3] ? (customer.orders[3].createdAt - customer.orders[2].createdAt) / (1000 * 60 * 60 * 24) : 'N/A';

    var row = [
      customer.customerId,
      customer.orders.length,
      firstOrderDate != 'N/A' ? Utilities.formatDate(firstOrderDate, Session.getScriptTimeZone(), "dd - MMM - yyyy") : 'N/A',
      secondOrderDate != 'N/A' ? Utilities.formatDate(secondOrderDate, Session.getScriptTimeZone(), "dd - MMM - yyyy") : 'N/A',
      thirdOrderDate != 'N/A' ? Utilities.formatDate(thirdOrderDate, Session.getScriptTimeZone(), "dd - MMM - yyyy") : 'N/A',
      fourthOrderDate != 'N/A' ? Utilities.formatDate(fourthOrderDate, Session.getScriptTimeZone(), "dd - MMM - yyyy") : 'N/A',
      avgTime1to2,
      avgTime2to3,
      avgTime3to4
    ];
    trendSheet.appendRow(row);
  }

  // Shopify <> Google Sheet Connector Developed by Abhishek Maity.
  // For more such micro tool contact abhishek.maity@adbuffs.com
  // Linkedin - https://www.linkedin.com/in/abhishek-maity-adbuffs/
}

