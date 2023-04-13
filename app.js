

const express = require('express')
const app = express()
const bodyparser = require('body-parser')
const fs = require('fs');
const readXlsxFile = require('read-excel-file/node');
const mysql = require('mysql')
const multer = require('multer')
const path = require('path')
//use express static folder
app.use(express.static("./public"))
// body-parser middleware use
app.use(bodyparser.json())
app.use(bodyparser.urlencoded({
    extended: true
}))
// Database connection
const db = mysql.createConnection({
    host: "localhost",
    user: "root",
    password: "abdullah",
    database: "readexcel1"
})
db.connect(function (err) {
    if (err) {
        return console.error('error: ' + err.message);
    }
    console.log('Connected to the MySQL server.');
})
// Multer Upload Storage
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, __basedir + '/uploads/')
    },
    filename: (req, file, cb) => {
        cb(null, file.fieldname + "-" + Date.now() + "-" + file.originalname)
    }
});
const upload = multer({ storage: storage });
//! Routes start
//route for Home page
app.get('/', (req, res) => {
    res.sendFile(__dirname + '/index.html');
});
// -> Express Upload RestAPIs
let __basedir = path.resolve();

app.post('/uploadfile', upload.single("uploadfile"), (req, res) => {
    if(!req.file.filename){
        return res.send('there is a problem')
    }
    importExcelData2MySQL(__basedir + '/uploads/' + req.file.filename);
    return res.sendFile(__dirname + '/thank.html');
});

// -> Import Excel Data to MySQL database
function importExcelData2MySQL(filePath) {
    // File path.
    readXlsxFile(filePath).then((rows) => {    
        rows.shift();
        // Open the MySQL connection
        const db = mysql.createConnection({
            host: 'localhost',
            user: 'root',
            password: 'abdullah',
            database: 'readexcel1'
        });
        //  CREATE TABLE `readexcel1`.`commissionscalculees` (
        //     `id` int NOT NULL AUTO_INCREMENT,
        //     `Code de production` bigint DEFAULT NULL,
        //     `Code adhérent` decimal(10,0) DEFAULT NULL,
        //     `Nom adhérent` varchar(45) DEFAULT NULL,
        //     `Type adhérent` varchar(45) DEFAULT NULL,
        //     `Garantie` varchar(255) DEFAULT NULL,
        //     `Période début` varchar(45) DEFAULT NULL,
        //     `Période fin` varchar(45) DEFAULT NULL,
        //     `Périodicité` varchar(45) DEFAULT NULL,
        //     `Type commission` varchar(255) DEFAULT NULL,
        //     `Prime HT` decimal(10,0) DEFAULT NULL,
        //     `Taux` int DEFAULT NULL,
        //     `Commission` decimal(10,0) DEFAULT NULL,
        //     `Date de calcul` varchar(455) DEFAULT NULL,
        //     `Entreprise` varchar(45) DEFAULT NULL,
        //     PRIMARY KEY (`id`)
        //   ) ENGINE=InnoDB AUTO_INCREMENT=317 DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_0900_ai_ci;
        db.connect((error) => {
            if (error) {
                console.error(error);
            } else {
                let    query="INSERT INTO `readexcel1`.`commissionscalculees` (`Code de production`,`Code adhérent`,`Nom adhérent`,`Type adhérent`,`Garantie`,`Période début`,`Période fin`,`Périodicité`,`Type commission`,`Prime HT`,`Taux`,`Commission`,`Date de calcul`,`Entreprise`) VALUES ?";
                // let query = 'INSERT INTO user (user_name, user_email) VALUES ?';
                db.query(query, [rows], (error, response) => {
                    console.log("🚀 ~ file: app.js:96 ~ db.query ~ error:", error)
                });
            }
        });
        
    })
}
// Create a Server
let server = app.listen(3000, function () {
    let host = server.address().address
    let port = server.address().port
    console.log("App listening at http://%s:%s", host, port)
})