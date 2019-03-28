var express =   require("express");
var multer  =   require('multer');
var XLSX = require('xlsx');


var mongoose     = require('mongoose');
var Schema       = mongoose.Schema;


var app         =   express();
var storage =   multer.diskStorage({
    destination: function (req, file, callback) {
        callback(null, './');
    },
    filename: function (req, file, callback) {
        callback(null, file.originalname);
    }
});
var upload = multer({ storage : storage}).single('userPhoto');

app.get('/',function(req,res){
    res.sendFile(__dirname + "/index.html");
});





app.post('/api/photo',function(req,res){
    upload(req,res,function(err) {
        if(err) {
            return res.end("Error uploading file.");
        }


//file upload code starts here


        var workbook = XLSX.readFile('practical_demo1.ods', {sheetStubs: true});
        var sheet_name_list = workbook.SheetNames;
        sheet_name_list.forEach(function(y) {
            var worksheet = workbook.Sheets[y];
            var headers = {};
            var data = [];
            var  myName  = [];
            var arr1=[];
            var arr2=[];


            //console.log(XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]], {blankCell : false}))

            for(z in worksheet) {
                if(z[0] === '!') continue;
                //parse out the column, row, and value
                var tt = 0;
                for (var i = 0; i < z.length; i++) {
                    if (!isNaN(z[i])) {
                        tt = i;
                        break;
                    }
                };
                var col = z.substring(0,tt);
                var row = parseInt(z.substring(tt));
                var value = worksheet[z].v;

                //store header names
                if(row == 1 && value) {
                    headers[col] = value;
               //     console.log(headers[col]);
                    continue;

                }

               // console.log('the following is the value of'+  row   +"=="+value);
                if(!data[row]) data[row]={};
                data[row][headers[col]] = value;
            }
            //drop those first two rows which are empty
            data.shift();
            data.shift();


            for (var i = 0; i < data.length; i++) {


                if(!data[i].hasOwnProperty("Priority")){
                    data[i].Priority='';
                   // myName.push(data[i]);

                    arr1.push(data[i]);
                }else{
                   // myName.splice(data[i].Priority-1, 0 , data[i]);

                    arr2.push(data[i]);
                }
            }

            var hello = arr1;
            arr2.sort((a, b) =>  parseFloat(a.Priority) - parseFloat(b.Priority)) ;


            for (var i = 0; i < arr2.length; i++) {

                //console.log("printing pro==",data[i].Priority)

                    hello.splice(arr2[i].Priority-1, 0 , arr2[i]);


            }



//Set up default mongoose connection
            var mongoDB = 'mongodb://localhost:27017/DEMO';
            mongoose.connect(mongoDB);
// Get Mongoose to use the global promise library
            mongoose.Promise = global.Promise;
//Get the default connection
            var db = mongoose.connection;

//Bind connection to error event (to get notification of connection errors)
            db.on('error', console.error.bind(console, 'MongoDB connection error:'))
//file data to mongodab

            var TmpSchema   = new Schema({
                Question: String,
                Option_1 : String,
                Option_2 : String,
                Option_3 : String,
                Option_4 : String,
                Answer   : String,
                Priority : Number

            });


            var Tmp= mongoose.model('Tmp', TmpSchema);
            module.exports = Tmp;



            res.end("File is uploaded");

            Tmp.insertMany(hello, function(error, docs) {});


        });

//file upload code ends here

    });
});

app.listen(3000,function(){
    console.log("Working on port 3000");
});