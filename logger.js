/**
 * Created by liudonghua on 2016-03-23.
 */
/** https://github.com/nomiddlename/log4js-node */
var log4js = require('log4js');
log4js.configure({
    appenders: [
        { type: 'console' },
        { type: 'file', filename: 'logs/app.log', category: 'app' }
    ]
});
var logger = log4js.getLogger('app');
logger.setLevel('INFO');

var getLogger = function() {
    return logger;
};

exports.logger = getLogger();