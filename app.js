var express = require('express');
var path = require('path');
var fs = require('fs');
var favicon = require('serve-favicon');
// use to log request info
var logger = require('morgan');
var cookieParser = require('cookie-parser');
var bodyParser = require('body-parser');
// see more on https://nodejs.org/api/util.html#util_util
var util = require('util');

// see more on https://github.com/nomiddlename/log4js-node
// another popular logging framework is winston (https://github.com/winstonjs/winston)
var logging = require('./logger.js').logger;

var resources_root = path.join(__dirname, 'resources');
var docxtemplater = require('docxtemplater');

// load the template file
var template = fs.readFileSync(path.join(resources_root, 'template', 'template.docx'), 'binary');

// parse excel data file
var xlsx = require('node-xlsx');
// see more on http://www.collectionsjs.com/fast-map
var FastMap = require("collections/fast-map");
var fastMap = FastMap();

/**
 * parse the xlsx file data and store in a memory map for quick search
 */
var parse_xlsx_file = function () {
    logging.info(util.format('parse excel data file'));
    var obj = xlsx.parse(path.join(resources_root, '/data.xlsx'));
    var data = obj[0].data;
    // skip header
    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (row.length == 7) {
            var format_row = {
                'ZKZH': row[0],
                'SFZH': row[1],
                'XM': row[2],
                'FSXY': row[3],
                'FSZY': row[4],
                'FSZYDM': row[5],
                'PYFS': row[6]
            };
            fastMap.add(format_row, format_row['ZKZH']);
        }
        else {
            logging.warn(util.format('illegal format record: %j', row));
        }
    }
    logging.info(util.format('finished parse excel data file with %d records', fastMap.length));
};

parse_xlsx_file();

var app = express();

// view engine setup
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// uncomment after placing your favicon in /public
app.use(favicon(path.join(__dirname, 'public', 'favicon.ico')));
app.use(logger('dev'));
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({extended: false}));
app.use(cookieParser());
app.use(express.static(path.join(__dirname, 'public')));

/**
 * check is valid user
 * @param detailed_info
 * @param examination_num
 * @param id_num
 * @param username
 * @returns {*|boolean}
 */
var is_valid_user = function (detailed_info, examination_num, id_num, username) {
    return detailed_info
        && detailed_info['ZKZH'] == examination_num
        && detailed_info['SFZH'] == id_num
        && detailed_info['XM'] == username;
};

// prepare generated or cached files
app.get("/:id_num/:examination_num/:username/:file", function (req, res) {
    var id_num = req.params["id_num"];
    var examination_num = req.params["examination_num"];
    var username = req.params["username"];
    var file = req.params["file"];
    logging.info(util.format('download file info-> id_num: %s, examination_num: %s, username: %s, file: %s', id_num, examination_num, username, file));
    // check is valid user
    var detailed_info = fastMap.get(examination_num);
    if (is_valid_user(detailed_info, examination_num, id_num, username)) {
        // extract examination number and check is valid
        var file_splited = file.split('.');
        if (file_splited.length != 2 && file_splited[0] != examination_num) {
            res.status(404);
            res.render('error', {
                message: "下载文件不存在",
                error: {}
            });
        }
        else {
            // check cache file
            var generated_file_path = path.join(resources_root, 'generated_notes', req.params["file"]);
            if (fs.existsSync(generated_file_path)) {
                logging.info(util.format('send cached download file : %s', generated_file_path));
                res.sendfile(generated_file_path);
            }
            else {
                generated_file_path = path.join(resources_root, 'generated_notes', examination_num + ".docx");
                var doc = new docxtemplater(template);
                var data = {
                    "name": detailed_info['XM'],
                    "college": detailed_info['FSXY'],
                    "major": detailed_info['FSZY'],
                    "major_code": detailed_info['FSZYDM'],
                    "cultivation": detailed_info['PYFS']
                };
                doc.setData(data);
                logging.info(util.format('fill template with data : %j', data));
                doc.render();
                var buf = doc.getZip().generate({type: "nodebuffer"});
                logging.info(util.format('create generated file : %s', generated_file_path));
                fs.writeFileSync(generated_file_path, buf);
                logging.info(util.format('finish create generated file : %s', generated_file_path));
                res.sendfile(generated_file_path);
            }
        }
    }
    else {
        res.render('info_invalid');
    }
});

// show basic info and download link
app.get("/:id_num/:examination_num/:username", function (req, res) {
    var id_num = req.params["id_num"];
    var examination_num = req.params["examination_num"];
    var username = req.params["username"];
    logging.info(util.format('logging info-> id_num: %s, examination_num: %s, username: %s', id_num, examination_num, username));
    // check if valid user
    var detailed_info = fastMap.get(examination_num);
    if (is_valid_user(detailed_info, examination_num, id_num, username)) {
        res.render('info', {
            id_num: detailed_info['SFZH'],
            examination_num: detailed_info['ZKZH'],
            username: detailed_info['XM'],
            college: detailed_info['FSXY'],
            major: detailed_info['FSZY'],
            major_code: detailed_info['FSZYDM'],
            cultivation: detailed_info['PYFS']
        });
    }
    else {
        res.render('info_invalid');
    }
});

// catch 404 and forward to error handler
app.use(function (req, res, next) {
    var err = new Error('你输入的信息不正确!');
    err.status = 404;
    next(err);
});

// error handlers

// development error handler
// will print stacktrace
if (app.get('env') === 'development') {
    app.use(function (err, req, res, next) {
        res.status(err.status || 500);
        res.render('error', {
            message: err.message,
            error: err
        });
    });
}

// production error handler
// no stacktraces leaked to user
app.use(function (err, req, res, next) {
    res.status(err.status || 500);
    res.render('error', {
        message: err.message,
        error: {}
    });
});


module.exports = app;
