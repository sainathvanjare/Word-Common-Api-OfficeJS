(function(root) {
    "use strict";

    if (!openXml)
        throw new Error("Failed to initiate, openxml sdk for javascript not loaded.");

    var XAttribute = Ltxml.XAttribute;
    var XCData = Ltxml.XCData;
    var XComment = Ltxml.XComment;
    var XContainer = Ltxml.XContainer;
    var XDeclaration = Ltxml.XDeclaration;
    var XDocument = Ltxml.XDocument;
    var XElement = Ltxml.XElement;
    var XEntity = Ltxml.XEntity;
    var XName = Ltxml.XName;
    var XNamespace = Ltxml.XNamespace;
    var XNode = Ltxml.XNode;
    var XObject = Ltxml.XObject;
    var XProcessingInstruction = Ltxml.XProcessingInstruction;
    var XText = Ltxml.XText;
    var cast = Ltxml.cast;
    var castInt = Ltxml.castInt;

    function getRelsPartUriOfPart(part) {
        var uri = part.uri;
        var lastSlash = uri.lastIndexOf('/');
        var partFileName = uri.substring(lastSlash + 1);
        var relsFileName = uri.substring(0, lastSlash) + "/_rels/" + partFileName + ".rels";
        return relsFileName;
    }

    function getRelsPartOfPart(part) {
        var relsFileName = getRelsPartUriOfPart(part);
        var relsPart = part.pkg.getPartByUri(relsFileName);
        return relsPart;
    }

    function addRelationshipToRelPart(part, relationshipId, relationshipType, target, targetMode) {
        var rxDoc = part.getXDocument();
        var tm = null;
        if (targetMode !== "Internal")
            tm = new XAttribute("TargetMode", "External");
        rxDoc.getRoot().add(
            new XElement(openXml.PKGREL.Relationship,
                new XAttribute("Id", relationshipId),
                new XAttribute("Type", relationshipType),
                new XAttribute("Target", target),
                tm));
    }

    function generateUUID(withhyphen) {
        var d = new Date().getTime();
        var format = 'xxxxxxxxxxxx4xxxyxxxxxxxxxxxxxxx';
        if (withhyphen)
            format = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx';
        var uuid = format.replace(/[xy]/g, function(c) {
            var r = (d + Math.random() * 16) % 16 | 0;
            d = Math.floor(d / 16);
            return (c == 'x' ? r : (r & 0x3 | 0x8)).toString(16);
        });
        return uuid;
    };

    openXml.OpenXmlPackage.prototype.addCustomXMLPart = function(data) {
        /*Parse and Validate Data*/
        var xmlDataEle = XElement.parse(data);

        /*Get the main part and find existing customXMLParts*/
        var mainPart = this.getPartByRelationshipType(openXml.relationshipTypes.mainDocument);
        var customXmlParts = mainPart.getPartsByRelationshipType(openXml.relationshipTypes.customXml);
        var index = (customXmlParts.length + 1);

        /*Create URI*/
        var baseFolder = "/customXml/";
        var itemFile = baseFolder + "item" + index + ".xml";
        var dataPartUri = itemFile;
        var itemPropsFile = "itemProps" + index + ".xml";
        var dataStorePartUri = baseFolder + itemPropsFile;

        /*Create Data store element and add parts*/
        var dataStoreEle = new XDocument(
            new XDeclaration("1.0", "utf-8", "yes"),
            new XElement(openXml.DS.datastoreItem,
                new XAttribute(XNamespace.xmlns + "ds", openXml.dsNs.namespaceName),
                new XAttribute(openXml.DS.itemID, "{" + generateUUID(true) + "}"),
                new XElement(openXml.DS.schemaRefs,
                    new XElement(openXml.DS.schemaRef, new XAttribute("uri", xmlDataEle.getDefaultNamespace())
                    ))));
        var newDataStorePart = new openXml.OpenXmlPart(this, dataStorePartUri, openXml.contentTypes.customXmlProperties, "xml", dataStoreEle);
        this.parts[dataStorePartUri] = newDataStorePart;
        var newPart = new openXml.OpenXmlPart(this, dataPartUri, openXml.contentTypes.customXmlProperties, "xml", xmlDataEle);
        this.parts[dataPartUri] = newPart;

        /*Updated Content_Type.xml*/
        this.ctXDoc.getRoot().add(
            new XElement(openXml.CT.Override,
                new XAttribute("PartName", dataStorePartUri),
                new XAttribute("ContentType", openXml.contentTypes.customXmlProperties)));

        /*Add relationships*/
        var relsPart = getRelsPartOfPart(newPart);
        if (!relsPart) {
            var relsPartUri = getRelsPartUriOfPart(newPart);
            relsPart = new openXml.OpenXmlPart(this, relsPartUri, openXml.contentTypes.customXmlProperties, "xml", new XDocument(
                new XDeclaration("1.0", "utf-8", "yes"),
                new XElement(openXml.PKGREL.Relationships,
                    new XAttribute(XNamespace.xmlns + "rel", openXml.pkgRelNs.namespaceName))));
            this.parts[relsPartUri] = relsPart;
        }
        var partRelId = "rId" + generateUUID(false);
        addRelationshipToRelPart(relsPart, partRelId, openXml.relationshipTypes.customXml, itemPropsFile, "Internal");
        var mainPartRelId = "rId" + generateUUID(false);
        mainPart.addRelationship(mainPartRelId, openXml.relationshipTypes.customXml, ".." + itemFile, "Internal");
    };

    openXml.OpenXmlPackage.prototype.getCustomXMLPartByNS = function(namespace, mainPart) {
        var parts = [];
        if (!mainPart)
            mainPart = this.getPartByRelationshipType(openXml.relationshipTypes.mainDocument);
        var customXmlParts = mainPart.getPartsByRelationshipType(openXml.relationshipTypes.customXml);
        if (namespace === "")
            namespace = "__none";
        for (var i = 0; i < customXmlParts.length; ++i) {
            if (customXmlParts[i].getXDocument().getRoot().getDefaultNamespace().namespaceName === namespace) {
                parts.push(customXmlParts[i]);
            }
        }
        return parts;
    };

} (this));