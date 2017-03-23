/**
 * Created by liudonghua on 2016-03-24.
 */
var path = require('path');
var fs = require('fs');
var util = require('util');

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

var generate_docx_file = function () {
    var hash_values = fastMap.values();
    for (var i = 0; i < hash_values.length; i++) {
        var detailed_info = hash_values[i];
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
        var generated_file_path = path.join(resources_root, 'generated_notes', detailed_info['ZKZH'] + ".docx");
        logging.info(util.format('create generated file : %s', generated_file_path));
        fs.writeFileSync(generated_file_path, buf);
        logging.info(util.format('finish create generated file : %s', generated_file_path));
    }
};

console.time('parse_xlsx_file');
parse_xlsx_file();
console.timeEnd('parse_xlsx_file');

console.time('generate_docx_file');
generate_docx_file();
console.timeEnd('generate_docx_file');