async = require('async');
JSZip = require("jszip");
var DOMParser = require('xmldom').DOMParser;

Enumerable = require('./resources/linq.js');
Ltxml = require('./resources/ltxml.js');
XEnumerable = require("./resources/ltxml-extensions.js");
openXml = require("./resources/openxml.js");
require("./resources/openxml-extension.js");

Ltxml.DOMParser = DOMParser;
module.exports = {
    openXml: openXml,
    Ltxml:Ltxml
}