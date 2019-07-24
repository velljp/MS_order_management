
var LINE_ITEM_NAME = "Lineitem name"
var ITEM_NUMBER = "商品番号"
var TAOBAO_LINK = "タオバオリンク"

function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu(
    'シート更新', 
    [
      {name: '上下反転', functionName: 'updown'},
      {name: 'インポート to result', functionName: 'main'}
    ]
  )
}

function updown(){
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
  var lastRow = activeSheet.getLastRow()
  var valuess = activeSheet.getDataRange().getValues()
  valuess.shift()
  var reverse = valuess.reverse()
  activeSheet.getRange(2, 1, reverse.length, reverse[0].length).setValues(reverse)
}

function extractHeader(sheet , headers){
  var headerRange = sheet.getRange("1:1")
  var lastCol = headerRange.getLastColumn()
  for(var i = 1 ; i <= lastCol ; i++){
    var v = headerRange.getCell(1, i).getValue()
    headers.map(function(colData){
      if(colData.key == v){colData.value = i}
    })
  }
  return headers
}


function setValues(extractSheet, orderSheet, headers){
  
  var lastRow = extractSheet.getLastRow()
  var margin = 200
  var startIndex = 2
  var lastIndex = Math.min(startIndex + margin, lastRow)
  while(startIndex <= lastRow){
    //なぜか１回のループで200~400回以上回せないようなので
    var valuess = []
    for(var i = startIndex; i <= lastIndex ; i++){
      var values = []
      headers.forEach(function(header){
        var v = extractSheet.getRange(i, header.value).getValue()
        values.push(v)
      })
      valuess.push(values)
    }
    
    resultSheet.getRange(startIndex, 1, valuess.length, valuess[0].length).setValues(valuess)

    startIndex = lastIndex + 1
    lastIndex = Math.min(startIndex + margin, lastRow)
  }    
}

function getHeaderIndex(sheet, headerStr){
  var headerRange = sheet.getRange("1:1")
  var lastCol = headerRange.getLastColumn()
  for(var i = 1; i <= lastCol ; i++){
    var header = headerRange.getCell(1, i).getValue()
    if(header == headerStr){return i}
  }
  return 0
}


function setOrderNumber(sheet, headers){
  
  var itemCol = getHeaderIndex(sheet, LINE_ITEM_NAME)
  var lastCol = headers.length
  
  var orderNums = []

  var lastRow = sheet.getLastRow()
  for(var i = 2; i <= lastRow ; i++){
    var v = sheet.getRange(i, itemCol).getValue()
    var result = v.match(/\s[A-Za-z0-9]+(\s|$)/g)
    if(result == null || result.length == 0){
      orderNums.push([""])
      continue
    }
    var orderNum = result[0].trim().toUpperCase()
    orderNums.push([orderNum])
  }
  
  sheet.getRange(2, lastCol+1, orderNums.length, 1).setValues(orderNums)
}

function setHeader(sheet, headers){
  for(var i = 0 ; i < headers.length ; i++){
    var header = headers[i]
    sheet.getRange(1, i+1).setValue(header.key)
  }
  sheet.getRange(1, headers.length+1).setValue(ITEM_NUMBER)
  sheet.getRange(1, headers.length+2).setValue(TAOBAO_LINK)
}


function setTaobaoLink(itemSheet, itemHeaders, orderSheet, orderHeaders){
  var orderLastRow = orderSheet.getLastRow()
  var itemLastRow = itemSheet.getLastRow()
  
  var itemNumIndexInOrder = orderHeaders.filter(function(v){
    return v.key == ITEM_NUMBER
  })[0].value
  var taobaoIndexInOrder = orderHeaders.filter(function(v){
    return v.key == TAOBAO_LINK
  })[0].value

  var itemNumIndexInItem = itemHeaders.filter(function(v){
    return v.key == ITEM_NUMBER
  })[0].value
  var taobaoIndexInItem = itemHeaders.filter(function(v){
    return v.key == TAOBAO_LINK
  })[0].value
  
  var taobaos = []
  
  var itemSheetData = itemSheet.getDataRange().getValues()
  var orderSheetData = orderSheet.getDataRange().getValues()
  
  for(var i = 1; i < orderSheetData.length; i++){
    var orderNum = orderSheetData[i][itemNumIndexInOrder-1]
    var taobao = ""    
    for(var j = 1; j < itemSheetData.length; j++){
      if(itemSheetData[j][itemNumIndexInItem-1] != orderNum){
        continue
      }
      taobao = itemSheetData[j][taobaoIndexInItem-1]
      break
    }    
    taobaos.push([taobao])
  }
  
  orderSheet.getRange(2, taobaoIndexInOrder, taobaos.length , 1).setValues(taobaos)
}


function main() {
  var importSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('インポート')
  var orderSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('result')
  var itemSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商品一覧')
  
  var extractHeaders = [
    {key:"Name", value:0}, 
    {key:"Paid at", value: 0}, 
    {key:"Billing Name", value: 0}, 
    {key:LINE_ITEM_NAME, value: 0}, 
    {key: "Id", value: 0},
  ]

  var itemHeaders = [
    {key:ITEM_NUMBER, value:0}, 
    {key:TAOBAO_LINK, value: 0}, 
  ]

  var orderHeaders = [
    {key:ITEM_NUMBER, value:0}, 
    {key:TAOBAO_LINK, value: 0}
  ]
  
  extractHeaders = extractHeader(importSheet, extractHeaders)
  setHeader(orderSheet, extractHeaders)
    
  setValues(importSheet, orderSheet, extractHeaders)  
  setOrderNumber(orderSheet, extractHeaders)

  itemHeaders = extractHeader(itemSheet, itemHeaders)
  orderHeaders = extractHeader(orderSheet, orderHeaders)
  setTaobaoLink(itemSheet, itemHeaders, orderSheet, orderHeaders)
  
  Browser.msgBox("MS／商品管理のスクリプトが終わりました")
}
