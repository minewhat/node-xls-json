var xlsjs = require('xlsjs');
var cvcsv = require('csv');

exports = module.exports = XLS_json;

// exports.XLS_json = XLS_json;

function XLS_json (config, callback,finalcallback) {
  if(!config.input) {
    console.error("You miss a input file");
  }

  var cv = new CV(config, callback,finalcallback);
  
}

function CV(config, callback) { 
  var wb = this.load_xls(config.input)
  var ws = this.ws(wb);
  var csv = this.csv(ws)
  this.cvjson(csv, config.records, callback,finalcallback)
}

CV.prototype.load_xls = function(input) {
  return xlsjs.readFile(input);
}

CV.prototype.ws = function(wb) {
  var target_sheet = '';

  if(target_sheet === '') 
    target_sheet = wb.SheetNames[0];
  ws = wb.Sheets[target_sheet];
  return ws;
}

CV.prototype.csv = function(ws) {
  return csv_file = xlsjs.utils.make_csv(ws)
}

CV.prototype.cvjson = function(csv, records , callback,finalcallback) {
  var record = []
  var header = []

  cvcsv()
    .from.string(csv)
    .transform( function(row){
      row.unshift(row.pop());
      return row;
    })
    .on('record', function(row, index){
      
      if(index === 0) {
        header = row;
      }else{
        var obj = {};
        header.forEach(function(column, index) {
          obj[column.trim()] = row[index].trim();
        })
        record.push(obj);
      }
      if( (index % records) == 0 )
      {
         callback(null,record);
         record = []
      }
    })
    .on('end', function(count){
      // when writing to a file, use the 'close' event
      // the 'end' event may fire before the file has been written
      callback(null, record);
      if(finalcallback)
        finalcallback(null);
      
    })
    .on('error', function(error){
      console.log(error.message);
      if(finalcallback)
        finalcallback(error.message);
    });
}
