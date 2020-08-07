// function to pull orders from the WooCommerce Rest API
function order_syncv2() {

  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_name = doc.getSheetByName('orders').getName();
  var sheet = doc.getSheetByName(sheet_name);
    
  fetch_orders(sheet_name);
  
  // to prevent duplicates in the data, I grab the Max of the created_date from my orders sheet and only pull data that is after that date 
  // may need to customized for your personal sheet setup
    
  var data = sheet.getRange(9,1,sheet.getLastRow() - 8,sheet.getLastColumn());
  data.sort({column: 6, ascending: true});
  
  var cell = sheet.getRange("B6");
  cell.setFormula("=MAX(F:F)");
    
}


// function to fetch orders from the WooCommerce API
function fetch_orders(sheet_name) {
  
  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = doc.getSheetByName(sheet_name);
    
  // this is how the sheet is setup -- there are better ways to do this, but for beginners, this is likely the easiest to understand 
  var ck = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B4").getValue();
  var cs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B5").getValue();
  var website = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B3").getValue();
  var manualDate = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getRange("B6").getValue(); // Set your order start date in spreadsheet in cell B6
  var m = new Date(manualDate).toISOString();
    
  // starts grabbing from page 1
  var page = 1;
  
  var surl = website + "/wp-json/wc/v2/orders?consumer_key=" + ck + "&consumer_secret=" + cs + "&after=" + m + "&per_page=100" + "&page=" + page; 
  
  var url = surl
  Logger.log(url)
  
  var options =
      {
          
        "method": "GET",
        "Content-Type": "application/x-www-form-urlencoded;charset=UTF-8",
        "muteHttpExceptions": true,
        
      };
  
  var result = UrlFetchApp.fetch(url, options);
  
  if (result.getResponseCode() == 200) {
    
    var params = JSON.parse(result.getContentText());
    
  }
  
  page++;
    
  // As long as data is returned, import the data to the sheet, and move on to the next page
  // GS has a 6 min max runtime, so you may need to start from a higher page at the beginning and run the script multiple times if you have a lot of orders to import at first
  while (params != "") {
    import_orders_to_sheet(sheet_name, params);
    
    var surl = website + "/wp-json/wc/v2/orders?consumer_key=" + ck + "&consumer_secret=" + cs + "&after=" + m + "&per_page=100" + "&page=" + page; 
    url = surl;
    result = UrlFetchApp.fetch(url, options);
    
    if (result.getResponseCode() == 200) {
      
      params = JSON.parse(result.getContentText());
      
      page++;
      
    }
    
  }
  
}


// Function to import orders from JSON API response from server
// The import is customized to my needs, and each implementation will likely be custom to the user
function import_orders_to_sheet(sheet_name, params) {

  var doc = SpreadsheetApp.getActiveSpreadsheet();
  var temp = doc.getSheetByName(sheet_name);
  var arrayLength = params.length;
    
  for (var i = 0; i < arrayLength; i++) {
    var a, c, e, d;
    var container = [];
      
    a = container.push(params[i]["id"]);
    a = container.push(params[i]["order_key"]);
    a = container.push(params[i]["created_via"]);
    a = container.push(params[i]["status"]);
    a = container.push(params[i]["currency"]);
    a = container.push(new Date(params[i]["date_created"] + "-0500"));
    a = container.push(new Date(params[i]["date_modified"] + "-0500"));
    a = container.push(params[i]["customer_id"]);
    a = container.push(params[i]["billing"]["first_name"]);
    a = container.push(params[i]["billing"]["last_name"]);
    a = container.push(params[i]["billing"]["address_1"]);
    a = container.push(params[i]["billing"]["address_2"]);
    a = container.push(params[i]["billing"]["city"]);
    a = container.push(params[i]["billing"]["state"]);
    a = container.push(params[i]["billing"]["postcode"]);
    a = container.push(params[i]["billing"]["phone"]);
    a = container.push(params[i]["billing"]["email"]);
    a = container.push(params[i]["shipping"]["first_name"] + " "+ params[i]["shipping"]["last_name"]+" "+ params[i]["shipping"]["address_2"]+" "+ params[i]["shipping"]["address_1"]+" "+params[i]["shipping"]["city"]+" "+params[i]["shipping"]["state"]+" "+params[i]["shipping"]["postcode"]+" "+params[i]["shipping"]["country"]); 
    a = container.push(params[i]["payment_method_title"]);
    a = container.push(params[i]["date_paid"]);
    a = container.push(params[i]["customer_note"]);
    
    // Loop to parse line_items from the results
    // to keep things simple, I just comma deliminate values when an order has mupltiple order lines
    // if you don't want to do this, you can look at the code from my subscription import to see how I create a new line for every line item
    c = params[i]["line_items"].length;
    
    var ids = "";
    var prods = "";
    var items = "";
    var qtys = "";
    var prices = "";
    var totals = "";
    var total_line_items_quantity = 0;
      
    for (var k = 0; k < c; k++) {
      var line_id, item, prod_id, qty, price, total;
      
      line_id = params[i]["line_items"][k]["id"];
      ids = ids + line_id + ", ";
      
      qty = params[i]["line_items"][k]["quantity"];
      qtys = qtys + qty + ", ";
      
      item = params[i]["line_items"][k]["name"];
      items = items + item + ", ";
      
      prod_id = params[i]["line_items"][k]["product_id"];
      prods = prods + prod_id + ", ";
      
      price = params[i]["line_items"][k]["price"];
      prices = prices + price + ", ";
      
      total = params[i]["line_items"][k]["total"];
      totals = totals + total + ", ";
      
      // loop to handle metadata within order line items
      e = params[i]["line_items"][k]["meta_data"].length;
      
      var meta_occupancy = "";
      var meta_cabintype = "";
      var meta_payment = "";
      var meta_partypass = "";
      var meta_roomtype = "";
      
      for (var s = 0; s < e; s++) {
        var keys, value;
          
        keys = params[i]["line_items"][k]["meta_data"][s]["key"];
        value = params[i]["line_items"][k]["meta_data"][s]["value"];
                
        if (keys == "occupancy") {
          var meta_occupancy = value + ", ";
        } else if (keys == "cabin-type") {
          var meta_cabintype = value + ", ";
        } else if (keys == "payment-type") {
          var meta_payment = value + ", ";
        } else if (keys == "party-pass") {
          var meta_partypass = value + ", ";
        } else if (keys == "room-type") {
          var meta_roomtype = value + ", ";
        }
        
      }
      
    }
    
    a = container.push(ids);
    a = container.push(prods);
    a = container.push(qtys);
    a = container.push(items);
    a = container.push(prices);
    a = container.push(totals);
    
    a = container.push(meta_occupancy);
    a = container.push(meta_cabintype);
    a = container.push(meta_payment);
    a = container.push(meta_partypass);
    a = container.push(meta_roomtype);
    
    a = container.push(params[i]["total"]); //Total price of all line items
    a = container.push(params[i]["discount_total"]); // Total discount applied to all line items
    
    
    // loop to parse metadata from the results
    e = params[i]["meta_data"].length;
    
    var meta_dob = "";
    var meta_gender = "";
    var meta_residency = "";
    var meta_roommate_name = "";
    var meta_subscription_renewal = "";
    var meta_loyalty = "";
    
    for (var s = 0; s < e; s++) {
      var keys, value;
        
      keys = params[i]["meta_data"][s]["key"];
      value = params[i]["meta_data"][s]["value"];
            
      if (keys == "date-of-birth1") {
        var meta_dob = value;
      } else if (keys == "gender1") {
        var meta_gender = value;
      } else if (keys == "state-of-residency1") {
        var meta_residency = value;
      } else if (keys == "roommate-name-s1") {
        var meta_roommate_name = value;
      } else if (keys == "_subscription_renewal") {
        var meta_subscription_renewal = value;
      } else if (keys == "vifp-loyalty1") {
        var meta_loyalty = value;
      }
      
    }
    
    a = container.push(meta_dob);
    a = container.push(meta_gender);
    a = container.push(meta_residency);
    a = container.push(meta_roommate_name);
    a = container.push(meta_subscription_renewal);
    a = container.push(meta_loyalty);
        
    // loop to parse coupon lines from the results
    e = params[i]["coupon_lines"].length;
    
    for (var s = 0; s < e; s++) {
      var coupon_code = params[i]["coupon_lines"][s]["code"];
      var coupon_discount = params[i]["coupon_lines"][s]["discount"];
      
    }
    
    a = container.push(coupon_code);
    a = container.push(coupon_discount);
    
    // push the data stored in the container to the sheet
    temp.appendRow(container);
    
  }
  
}
