/*
    Copyright 2010 by Mathias Mamsch
    This file is part of the DOORS Standard Library 

    The DOORS Standard Library  is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    The DOORS Standard Library  is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with the DOORS Standard Library.  If not, see <http://www.gnu.org/licenses/>.
*/

/*! \file XMLDocument.inc 
\brief This file provides a class for reading, writing and manipulating XML data using the MSXML DOM Document class.
 
This library provides a wrapper for the MSXML DOM Document COM interface. For this it supplies three new datatypes:
- XMLDocument   => Class for an XML document. The class provides functions for manipulating XML tags or xslt transformations.
- XMLTag        => Class for the content of one xml tag. Provides functions for reading the content, reading the tag hierarchy.
- XMLNodelist   => Class for a list of nodes that is returned by the \em selectNodes function or the \em childNodes function.
- XMLSchemaCache => Class for the schema cache property of an XML document

\class XMLDocument

The XMLDocument object contains of a hierarchy of XMLTags. You can create an XMLDocument object by using the create_XMLDocument() 
function and delete it again by calling delete_XMLDOcument(XMLDocument &). The following general example shows how easy it is to 
generate an XML document using this library: 

\include general_xml_example.dxl

This example will produce the following XML: 
\verbatim
<Phonebook>
    <Person Gender="male">
        <Name>Mr. Schmidt</Name>
        <Phone>012-3456</Phone>
    </Person>    
    <Person Gender="female">
        <Name>Mrs. Meyer</Name>
        <Phone>056-7890</Phone>
    </Person>
</Phonebook>
\endverbatim

\class XMLTag
\class XMLNodelist
\class XMLSchemaCache
*/

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// Constants: Error Handling
//////////////////////////////////////////////////////////////////////////////////////////////////////////
const int XMLNOERR     =  0;
const int XMLLOADERR   = -1;
const int XMLPARSERR   = -2;
const int XMLVALERR    = -3;
const int XMLXFORMERR  = -4;
const int XMLSAVERR    = -5;
const int XMLNOFILEERR = -6;
const int XMLINVALERR  = -7;

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// Helper Functions: XML file management
//////////////////////////////////////////////////////////////////////////////////////////////////////////

//=================================================================
// Function:    fileSize()
// Usage:   Tests if a file exists and returns file size
//
// Input:   Filename
//
// Output:  File size in [byte]
//          -1: File not found
//=================================================================
int fileSize(string filename)
{
    int fileSize = -1
    Stat s
    s = create(filename)
    if (!null s) {
        fileSize = size(s)
        delete s
    }
    return fileSize
}

//=================================================================
// Function:    checkURL()
// Usage:   Tests if a passed URL is valid and points to an XML
//          file, that can be downloaded.
//
// Input:   URL
//
// Output:  True if the URL points to an XML file with read access;
//          otherwise False
//=================================================================
bool checkURL(string pURL)
{
    bool retVal = false

    OleAutoObj oleHttpRequest = oleCreateAutoObject("Msxml2.XMLHTTP.6.0")
    OleAutoObj oleDoc = null
    OleAutoObj oleDocRoot = null
    int iHttpStatus = 0

    if (!null oleHttpRequest) {
        oleMethod (oleHttpRequest, "open", oleArgs <- "GET" <- pURL <- false)
        oleMethod (oleHttpRequest, "send")
        oleGet (oleHttpRequest, "status", iHttpStatus)
        if (iHttpStatus == 200) {
            oleGet (oleHttpRequest, "responseXML", oleDoc)
            if (!null oleDoc) {
                oleGet (oleDoc, "documentElement", oleDocRoot)
                if (!null oleDocRoot) {
                    retVal = true
                    oleCloseAutoObject oleDocRoot
                }
                else {
                    // print "Empty or no XML document\n"
                }
                oleCloseAutoObject oleDoc
            }
        }
        else {
            // print "HTML error: " iHttpStatus "\n"
        }
        oleCloseAutoObject oleHttpRequest
    }

    return retVal
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// Helper Functions: XML general
//////////////////////////////////////////////////////////////////////////////////////////////////////////

//=================================================================
// Function:    create_XMLElementName()
// Usage:   Transforms a passed string into a valid XML element name
//
// Input:   Porposed XML element name
//
// Output:  Validated XML element name
//
//=================================================================
Buffer create_XMLElementName(Buffer &p_original)
{
    Regexp GET_INVALID_XML_CHAR = regexp2 "([^A-Za-z0-9_\\.]+)"
    Regexp HAS_INVALID_XML_START = regexp2 "^([xX][mM][lL]|[^A-Za-z_])"
    string pReplaceWith = "_"
    Buffer bTmp = p_original
    Buffer ret_buf

    if (!null bTmp )
    {
        string_replace(GET_INVALID_XML_CHAR, pReplaceWith, p_original)
    }
    if (null bTmp ||
        length(bTmp) == 0 ||
        HAS_INVALID_XML_START p_original)
    {
        ret_buf = create
        ret_buf += "_"
        if (!null bTmp)
        {
            combine(ret_buf, p_original, 0)
        }
        delete p_original
        p_original = ret_buf
    }

    ret_buf = p_original

    delete GET_INVALID_XML_CHAR
    delete HAS_INVALID_XML_START

    return ret_buf
}

//=================================================================
// Function:    create_XMLElementName()
// Usage:   Transforms a passed string into a valid XML element name
//
// Input:   Porposed XML element name
//
// Output:  Validated XML element name
//
//=================================================================
string create_XMLElementName(string p_original)
{
    string XMLElementName
    Buffer bTmp = create
    bTmp += p_original
    bTmp = create_XMLElementName(bTmp)
    XMLElementName = stringOf bTmp
    delete bTmp
    return XMLElementName
}

bool xmlNoValue ( string s ) { 
    if (s[0:7] == "Method '") return true
    return false
}

void xmlCheckOLE( string s) {    
    if (!null s && !xmlNoValue s) raiseError ("XML OLE Error in " (getCaller 3) "\ncalled from: " (getCaller 4) "\n", s)    
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////
// XML DOM API
//////////////////////////////////////////////////////////////////////////////////////////////////////////

//! \brief XMLDocument is a class that encapsules an XML document.
struct XMLDocument {};
DxlObject DxlObjectOf (XMLDocument xd) { return (addr_ xd) DxlObject }
XMLDocument XMLDocumentOf (DxlObject dx) { return (addr_ dx) XMLDocument }

// OLE Handle Properties
OleAutoObj getOleHandle_ (XMLDocument xd)                    { DxlObject x = DxlObjectOf xd; return (x->"oleHandle") OleAutoObj }
void       setOleHandle_ (XMLDocument xd, OleAutoObj oleObj) { DxlObject x = DxlObjectOf xd;         x->"oleHandle" = oleObj    }

//! \brief XMLTag is a class that encapsules an XML tag in an XML document.
struct XMLTag {};
DxlObject DxlObjectOf (XMLTag xd) { return (addr_ xd) DxlObject }

// OLE Handle Properties
OleAutoObj getOleHandle_ (XMLTag xt)                    { DxlObject x = DxlObjectOf xt; return (x->"oleHandle") OleAutoObj }
void       setOleHandle_ (XMLTag xt, OleAutoObj oleObj) { DxlObject x = DxlObjectOf xt;         x->"oleHandle" = oleObj    }

//! \brief XMLSchemaCache is a class that encapsules an XML schema cache.
struct XMLSchemaCache {};
DxlObject DxlObjectOf (XMLSchemaCache xs) { return (addr_ xs) DxlObject }
XMLSchemaCache XMLSchemaCacheOf (DxlObject dx) { return (addr_ dx) XMLSchemaCache }

// OLE Handle Properties
OleAutoObj getOleHandle_ (XMLSchemaCache xs)                    { DxlObject x = DxlObjectOf xs; return (x->"oleHandle") OleAutoObj }
void       setOleHandle_ (XMLSchemaCache xs, OleAutoObj oleObj) { DxlObject x = DxlObjectOf xs;         x->"oleHandle" = oleObj    }

XMLSchemaCache getSchemaCache_ (XMLDocument xd)                     { DxlObject x = DxlObjectOf xd; return (x->"schemaCache") XMLSchemaCache }
void           setSchemaCache_ (XMLDocument xd, XMLSchemaCache xs ) { DxlObject x = DxlObjectOf xd;         x->"schemaCache" = xs            }

/*! \memberof XMLSchemaCache
    \param xsc XMLSchemaCache object to be destroyed. Will be set to NULL by the function.
    \brief This function will destroy the XML schema cache

    This function must be called for every XML schema cache that was created using create_XMLSchemaCache_()
*/
void delete_XMLSchemaCache_ (XMLSchemaCache &xsc) {

    OleAutoObj oleObj = getOleHandle_ xsc; oleCloseAutoObject oleObj
    DxlObject x = DxlObjectOf xsc; delete x;
    xsc = null;
}

/*! \memberof XMLTag
    \brief returns the XMLDocument that contains the XMLTag xt
    \param xt The XMLTag whose parent to return
    \returns The XMLDocument that contains the XMLTag xt
*/
XMLDocument getParentDocument   (XMLTag xt)                   { DxlObject x = DxlObjectOf xt; return (x->"parentDoc") XMLDocument }
void        setParentDocument_  (XMLTag xt, XMLDocument xd  ) { DxlObject x = DxlObjectOf xt;         x->"parentDoc" = xd         }

Skip        getTagCache_   (XMLDocument xd)           { DxlObject x = DxlObjectOf xd; return (x->"tagCache") Skip  }
void        setTagCache_   (XMLDocument xd, Skip sk ) { DxlObject x = DxlObjectOf xd;         x->"tagCache" = sk   }

void cacheXMLTag_        (XMLDocument xd, XMLTag xt) { put (getTagCache_ xd, xt, xt, true) }
void removeCachedXMLTag_ (XMLDocument xd, XMLTag xt) { delete (getTagCache_ xd, xt) }

/*! \memberof XMLTag
    \param xd XMLTag object to be destroyed. Will be set to NULL by the function.
    \brief This function will destroy the XMLTag in memory and release its OLE handle.

    This function is called by delete_XMLDocument(XMLDocument &) which will delete all tags created in it. So
    you usually don't need to call it, except if you are handling huge documents and want to take care of memory.
*/
void delete_XMLTag (XMLTag &xt) {
    removeCachedXMLTag_ (getParentDocument xt, xt)
    OleAutoObj oleObj = getOleHandle_ xt; oleCloseAutoObject oleObj
    DxlObject x = DxlObjectOf xt; delete x;
    xt = null;
}


/*! \memberof XMLDocument
    \return Buffer with the formatted parsing error information. Null if parsing was successful.
    \param xd The XMLDocument that has been loaded (and parsed).
    \brief This function returns information about the first parsing error of an XML document object.
*/
Buffer getParseErrorMsg_ (XMLDocument xd) {
    // retrieve XML parse error
    OleAutoObj oleParseError
    Buffer getParseErrorMsg = null
    string sResult = ""
    int iResult = 0

    checkOleSuccess (  oleGet (getOleHandle_ xd, "parseError", oleParseError)  )
    checkOleSuccess (  oleGet (oleParseError, "errorCode", iResult )  )
    if (iResult != 0)
    {
        Regexp FIND_NS_URI = regexp2 "{[^}]+}"
        getParseErrorMsg = create
        getParseErrorMsg += "Error #1 in line: "
        checkOleSuccess (  oleGet (oleParseError, "line", iResult )  )
        getParseErrorMsg += iResult ""
        getParseErrorMsg += ", position: "
        checkOleSuccess (  oleGet (oleParseError, "linepos", iResult )  )
        getParseErrorMsg += iResult ""
        getParseErrorMsg += "\n"
        checkOleSuccess (  oleGet (oleParseError, "reason", sResult )  )
        getParseErrorMsg += sResult
        getParseErrorMsg += "Source: "
        checkOleSuccess (  oleGet (oleParseError, "srcText", sResult )  )
        getParseErrorMsg += replace(trimWhitespace(sResult), "\t", " ")
        getParseErrorMsg += "\n"

        string_replace(FIND_NS_URI, "", getParseErrorMsg)
        delete FIND_NS_URI
    }

    return getParseErrorMsg
}

/*! \memberof XMLDocument
    \return Buffer with the formatted validation error information. Null if validation was successful.
    \param[in] xd The XMLDocument that has been validated
    \param[in,out] error_buffer Buffer where to store the validation errors
    \brief This function returns information about the validation errors of an XML document object.
    \details If the validation directly follows a failed parse of the XML document,
             the parse error message should be passed as error_buffer. This function
             evaluates the parse error and updates the buffer with the found validation
             errors.
*/
Buffer getValidationErrorMsg_ (XMLDocument xd, Buffer &error_buffer) {
    // retrieve XML parse error
    const int XML_E_NODTD         = -1072897500
    const int XML_E_NOTWF         = -1072897501
    const int XML_INVALID_CONTENT = -1072898028

    Buffer getValidationErrorMsg = error_buffer
    if (null getValidationErrorMsg)
    {
        getValidationErrorMsg = create
    }

    OleAutoObj oleValidationError = null
    OleAutoObj oleValidationErrorList = null
    string sResult = ""
    int iResult = 0
    int iNumErrors = 0
    int i = 0

    if (checkOleSuccess (  oleMethod (getOleHandle_ xd, "validate", oleArgs(), oleValidationError)  ))
    {
        checkOleSuccess (  oleGet (oleValidationError, "errorCode", iResult )  )
        if (iResult != 0)
        {
            if (iResult == XML_E_NOTWF)
            {/*
              * The validation was interrupted. This is usually the case if
              * the xml file contains invalid characters.
              * Only the first validation error that was found during parsing
              * will be returned.
              */
              getValidationErrorMsg += "ERROR: Failed to completely parse and validate XML\n"
            }
            else
            {
                Regexp FIND_NS_URI = regexp2 "{[^}]+}"
                Regexp FIND_LN_INFO = regexp2 "/\\*\\[local-name\\(\\)="
                Regexp FIND_NS_INFO = regexp2 " and namespace-uri\\(\\)=\"[^\"]+\"\\]"
                checkOleSuccess (  oleGet (oleValidationError, "allErrors", oleValidationErrorList )  )
                if (!null oleValidationErrorList)
                {
                    checkOleSuccess (  oleGet (oleValidationErrorList, "length", iNumErrors )  )
                }

                if (iResult == XML_E_NODTD)
                {   /*
                    * This is a special case where parsing fails but
                    * validation after loading succeeds. Error information
                    * for the first error has now also to be taken from the validation.
                    */
                    setempty getValidationErrorMsg
                    getValidationErrorMsg += "Error #"
                    getValidationErrorMsg += ++i ""
                    getValidationErrorMsg += " of "
                    getValidationErrorMsg += iNumErrors ""
                    getValidationErrorMsg += ":\n"
                    checkOleSuccess (  oleGet (oleValidationError, "reason", sResult )  )
                    getValidationErrorMsg += sResult
                    getValidationErrorMsg += "Source: "
                    checkOleSuccess (  oleGet (oleValidationError, "errorXPath", sResult )  )
                    getValidationErrorMsg += sResult
                    getValidationErrorMsg += "\n\n"
                }

                if (iNumErrors > 1)
                {
                    if (i == 0)
                    {
                        setempty getValidationErrorMsg
                    }
                    checkOleSuccess (  oleGet (oleValidationErrorList, "next", oleValidationError )  )
                    while (!null oleValidationError)
                    {
                        getValidationErrorMsg += "Error #"
                        getValidationErrorMsg += ++i ""
                        getValidationErrorMsg += " of "
                        getValidationErrorMsg += iNumErrors ""
                        getValidationErrorMsg += ":\n"
                        checkOleSuccess (  oleGet (oleValidationError, "reason", sResult )  )
                        getValidationErrorMsg += sResult
                        getValidationErrorMsg += "Source: "
                        checkOleSuccess (  oleGet (oleValidationError, "errorXPath", sResult )  )
                        getValidationErrorMsg += sResult
                        getValidationErrorMsg += "\n\n"
                        checkOleSuccess (  oleGet (oleValidationErrorList, "next", oleValidationError )  )
                    }
                }
                string_replace(FIND_LN_INFO, " ", getValidationErrorMsg)
                string_replace(FIND_NS_INFO, "", getValidationErrorMsg)
                string_replace(FIND_NS_URI, "", getValidationErrorMsg)
                delete FIND_NS_URI
                delete FIND_LN_INFO
                delete FIND_NS_INFO
            }
        }
    }
    else
    {
        getValidationErrorMsg += "ERROR: Cannot validate XML\n"
    }

    return getValidationErrorMsg
}

/*! \memberof XMLDocument
    \return Buffer with the formatted validation error information. Null if validation was successful.
    \param[in] xd The XMLDocument that has been validated
    \brief This function returns information about the validation errors of an XML document object.
*/
Buffer getValidationErrorMsg_ (XMLDocument xd) {
    Buffer getValidationErrorMsg = create

    return getValidationErrorMsg_ (xd, getValidationErrorMsg)
}

/*! \memberof XMLDocument
    \return error code; 0 - if the XML document could be loaded and parsed correctly.
    \param xd The XMLDocument into which the contents shall be loaded.
    \param pXMLFile The XML file to load. Can be a file system path or an URL
    \brief This function loads the XML content of a file into an XMLDocument object.
*/
int loadXMLfile (XMLDocument xd, string pXMLFile) {
    int iRetVal = 0
    bool result = false
    bool validLocation = false

    if (fileSize( pXMLFile ) > 0)
    {
        validLocation = true
    }
    else if (checkURL( pXMLFile ))
    {
        /*
         * We do not want to setup an onreadystatechange handler to check
         * for complete download and parsing of the XML file
         */
        checkOleSuccess ( olePut (getOleHandle_ xd, "async", false) )
        validLocation = true
    }

    if (validLocation)
    {
        if (checkOleSuccess(  oleMethod (getOleHandle_ xd, "load", oleArgs <- pXMLFile, result)  ))
        {
            Buffer bError = null
            Buffer bParseMsg = getParseErrorMsg_ xd
            if (!null bParseMsg)
            {
                bError = create
                bError += "XMLPARSERR for "
                bError += "<"
                bError += pXMLFile
                bError += ">:\n"
                bError += "Parsing of XML file failed with following error(s):\n"

                /*
                 * We have encountered errors during parsing of the XML file.
                 * Reload and perform a full validation afterwards, if possible:
                 */
                checkOleSuccess ( olePut (getOleHandle_ xd, "validateOnParse", false) )
                if (checkOleSuccess(  oleMethod (getOleHandle_ xd, "load", oleArgs <- pXMLFile)  ))
                {
                    bParseMsg = getValidationErrorMsg_(xd, bParseMsg)
                }

                combine(bError, bParseMsg, 0)
                print tempStringOf bError

                iRetVal = XMLPARSERR
            }
            else
            {
                iRetVal = 0
            }
            delete bParseMsg
            delete bError
        }
        else
        {
            raiseError( "XMLLOADERR for " pXMLFile ":", "Cannot load XML file"  )
            iRetVal = XMLLOADERR
        }
    }
    else
    {
        raiseError( "XMLNOFILEERR for " pXMLFile ":", "Cannot open XML file"  )
        iRetVal = XMLNOFILEERR
    }
    return iRetVal
}

/*! \memberof XMLDocument
    \return true - if the XML document could be loaded and parsed correctly. 'false' if there was an error loading the document.
    \param xd The XMLDocument into which the contents shall be loaded.
    \param pXMLFile the file name of the XML document to load
    \brief This function loads the XML content of a file into an XMLDocument object.
*/
bool loadXMLfile (XMLDocument xd, string pXMLFile) {
    bool result = true
    int iRetVal = loadXMLfile (xd, pXMLFile)
    if (iRetVal < 0) result = false
    return result
}

/*! \memberof XMLDocument
    \return true - if the XML document could be loaded and parsed correctly. 'false' if there was an error loading the document.
    \param xd The XMLDocument into which the contents shall be loaded.
    \param s the contents of the XML document to load
    \brief This function loads the XML content of a string into an XMLDocument object.
*/
bool loadXML (XMLDocument xd, string s)
{
    bool result = false

    if (checkOleSuccess( oleMethod( getOleHandle_ xd, "loadXML", oleArgs <- s, result ) ))
    {
        if (!null s)
        {
            Buffer bErrMsg = null
            bErrMsg = getParseErrorMsg_ xd
            if (!null bErrMsg)
            {

                /*
                 * We have encountered errors during parsing of the XML document.
                 * Reload and perform a full validation afterwards, if possible:
                 */
                checkOleSuccess ( olePut (getOleHandle_ xd, "validateOnParse", false) )
                if (checkOleSuccess(  oleMethod (getOleHandle_ xd, "loadXML", oleArgs <- s)  ))
                {
                    bErrMsg = getValidationErrorMsg_(xd, bErrMsg)
                }

                raiseError( "XMLPARSERR",
                            "Parsing of XML document failed with following error(s):\n"         //-
                            tempStringOf(bErrMsg) )
            }
            else
            {
                result = true
            }
            delete bErrMsg
        }
    }
    return result
}

/*! \memberof XMLDocument
    \brief Same as loadXML(XMLDocumet, string) just takes a buffer as an argument
*/
bool loadXML (XMLDocument xd, Buffer s)
{
    bool result = false

    if (checkOleSuccess( oleMethod( getOleHandle_ xd, "loadXML", oleArgs <- s, result ) ))
    {
        if (length s > 0)
        {
            Buffer bErrMsg = null
            bErrMsg = getParseErrorMsg_ xd
            if (!null bErrMsg)
            {

                /*
                 * We have encountered errors during parsing of the XML document.
                 * Reload and perform a full validation afterwards, if possible:
                 */
                checkOleSuccess ( olePut (getOleHandle_ xd, "validateOnParse", false) )
                if (checkOleSuccess(  oleMethod (getOleHandle_ xd, "loadXML", oleArgs <- s)  ))
                {
                    bErrMsg = getValidationErrorMsg_(xd, bErrMsg)
                }

                raiseError( "XMLPARSERR",
                            "Parsing of XML document failed with following error(s):\n"         //-
                            tempStringOf(bErrMsg) )
            }
            else
            {
                result = true
            }
            delete bErrMsg
        }
    }
    return result
}

int save_XMLDocument_ (XMLDocument xd, string pFileName) {
    int iResult = 0
    string res = oleMethod (getOleHandle_ xd, "Save", oleArgs <- pFileName)
    if (!null res) {
        raiseError( "XMLSAVERR for XML file " pFileName ":",
                    "Saving XML document failed with following error:\n"         //-
                    res )
        iResult = XMLSAVERR
    }
    return iResult
}

/*! \memberof XMLDocument
    \return (Empty) XML document; returns null if XMLDocument object could not be created
    \brief Constructor for XMLDocument class
    This function returns a new and empty XML Document.
*/
XMLDocument create_XMLDocument () {
    OleAutoObj oleDoc = oleCreateAutoObject("Msxml2.DOMDocument.6.0")

    if (null oleDoc) {
        raiseError ("COM Error", "Microsoft XML DOMDocument seems not to be installed. Install Internet Explorer.")
        return null XMLDocument
    }

    // Parse complete XML document
    checkOleSuccess ( olePut (oleDoc, "validateOnParse", true) )
    checkOleSuccess ( oleMethod (oleDoc, "setProperty", oleArgs <- "MultipleErrorMessages" <- true ) )
    checkOleSuccess ( oleMethod (oleDoc, "setProperty", oleArgs <- "ResolveExternals" <- true ) )

    DxlObject x = new();
    XMLDocument result = (addr_ x) XMLDocument

    setOleHandle_ (result, oleDoc)
    Skip sk = create(); setTagCache_ (result, sk)

    loadXML (result, "")

    return result
}

void cleanupTags_ (XMLDocument xd) {
    // delete tag cache
    XMLTag xt
    Skip sk = getTagCache_ xd
    for xt in sk do delete_XMLTag xt
    delete sk
    setTagCache_ (xd, create Skip) 
}

/*! \memberof XMLDocument
    \param xd XMLDocument object to be destroyed. Will be set to NULL by the function.
    \brief This function will destroy the XMLDocument and all XMLTag objects that have been created from it in memory and release 
           their OLE handles. 
    
    This function must be called for every XMLDocument that was created using create_XMLDocument()
*/  
void delete_XMLDocument (XMLDocument &xd) {
    XMLDocument xdTemp = xd
    if (!null xdTemp)
    {
        cleanupTags_ xdTemp
        Skip sk = getTagCache_ xdTemp; delete sk

        XMLSchemaCache xsc = getSchemaCache_ xdTemp
        if (!null xsc)
        {
            delete_XMLSchemaCache_ xsc
        }

        OleAutoObj oleObj = getOleHandle_ xdTemp; oleCloseAutoObject oleObj
        DxlObject x = DxlObjectOf xdTemp; delete x;
        xd = xdTemp = null;
    }
}


/*! \memberof XMLSchemaCache
    \return (Empty) XML schema cache; returns null if XMLSchemaCache object could not be created
    \brief Constructor for XMLSchemaCache class
    This function returns a new and empty XML Schema Cache.
*/
XMLSchemaCache create_XMLSchemaCache_ () {
    OleAutoObj oleCache = oleCreateAutoObject("Msxml2.XMLSchemaCache.6.0")

    if (null oleCache) {
        raiseError ("COM Error", "Microsoft XML DOMDocument seems not to be installed. Install Internet Explorer.")
        return null XMLSchemaCache
    }

    checkOleSuccess ( olePut (oleCache, "validateOnLoad", true) )

    DxlObject x = new();
    XMLSchemaCache result = (addr_ x) XMLSchemaCache

    setOleHandle_ (result, oleCache)

    return result
}

/*! \memberof XMLSchemaCache
    \internal
    \return Error code in case of failure; otherwise false
    \param xmlns The namespace to associate with the passed XML schema.
    \param xd XML document object, that contains the XML schema
    \brief This function will add the XML schema to the schema cache.
    \details If a null XML document is passed, any schema for the passed
             namespace is removed from the schema cache.
*/
int   addXMLschema_ (XMLSchemaCache xc, string xmlns, XMLDocument xd) {

    if (null xd) {
        raiseError ("NULL Parameter", "XML Schema Document");
        return XMLINVALERR
    }

    OleAutoObj objCache = getOleHandle_ xc
    if (null objCache) {
        raiseError ("NULL Parameter", "XML Schema Cache");
        return XMLINVALERR
    }

    if (!checkOleSuccess ( oleMethod (objCache, "add", oleArgs <- xmlns <- (getOleHandle_ xd) ) )) {
        return XMLVALERR
    }

    return XMLNOERR
}


XMLTag create_XMLTag_ (XMLDocument xd, OleAutoObj x) {
    DxlObject dxlObj = new(); XMLTag result = (addr_ dxlObj) XMLTag
    setOleHandle_ (result, x) 
    setParentDocument_ (result, xd) 
    cacheXMLTag_ (xd, result)
    return result
}

/*! \memberof XMLDocument 
    \return A newly XMLTag object, that is associated to the XMLDocument xd. 
    \param xd The XMLDocument to associate with the newly created XMLTag
    \param tagName The name of the tag to create.
    \brief Creates a new XMLTag but does NOT put it to the hierarchy (See appendChild(XMLDocument, XMLTag))
    
    You need to call \b appendChild or \b insertChildBefore to make the new XMLTag show up in the document.
*/  
XMLTag create_XMLTag (XMLDocument xd, string tagName) {    
    if (null xd) { raiseError ("NULL Parameter", "XML Document"); return null XMLTag}
    
    OleAutoObj objDoc = getOleHandle_ xd
    if (null objDoc) { raiseError ("NULL Parameter", "XML Document"); return null XMLTag }
    
    OleAutoObj oleNewTag = null
    checkOLE ( oleMethod (objDoc, "createElement", oleArgs <- tagName, oleNewTag) )    
    if (null oleNewTag) { raiseError ("XML Error", "Could not create new element " tagName " in XML document."); }    
    
    return create_XMLTag_ (xd, oleNewTag) 
}


/*! \memberof XMLDocument
    \return error code; 0 - if the XML schema could be loaded and attached correctly.
    \param xmlns The namespace to associate with the XML schema.
    \param[in,out] pFileName File name of the XML schema file to load. Returns a
                             valid URI to the XML schema, if the schema has been
                             successfully attached.
    \note The returned URI must be used as schema location in new XML documents.
          If schema location in the XML file differs from the URI, that MSXML
          has stored in the schema cache, validation of the XML document will fail
          with error: The node is neither valid nor invalid because no DTD/Schema
          declaration was found.
    \brief This function will attach an XML schema to the XML document.
*/
int attachXMLschema (XMLDocument xd, string xmlns, string &pFileName) {
    XMLSchemaCache xsc = getSchemaCache_ xd
    XMLDocument xs = null
    string pUri = ""
    string tmpFileName = pFileName
    int iRetVal = 0

    if (null tmpFileName)
    {   // In MSXML 6.0 a null filename requests the removal of any schema for the
        // passed namespaces. This feature is not supported via OLE interface.
        raiseError( "NULL Parameter", "XML Schema file"  )
        iRetVal = XMLNOFILEERR
        return iRetVal
    }

    xs = create_XMLDocument
    iRetVal = loadXMLfile(xs, pFileName)
    if (iRetVal == 0)
    {
        if (null xsc)
        {
            xsc = create_XMLSchemaCache_
            setSchemaCache_ (xd, xsc)
            checkOleSuccess ( olePut ( getOleHandle_ xd, "schemas", getOleHandle_ xsc ) )
        }
        iRetVal = addXMLschema_(xsc, xmlns, xs)
    }

    if (iRetVal == 0)
    {
        checkOleSuccess (  oleGet (getOleHandle_ xs, "url", pUri)   )
        if (!null pUri) {
            pFileName = pUri
        }
    }

    delete_XMLDocument xs

    return iRetVal
}


void appendChildNode_ (DxlObject dxlParent, XMLTag xtChild) {
    if (null dxlParent) { raiseError ("NULL Parameter", "parent "); return}
        
    // cast to XMLDocument -> make sure that all items use the same property name for oleHandle
    XMLDocument xd = (addr_ dxlParent) XMLDocument
    
    OleAutoObj oleParent = getOleHandle_ xd
    OleAutoObj oleChild  = getOleHandle_ xtChild
    OleAutoObj oleDomElement

    checkOleSuccess ( oleMethod(oleParent, "appendChild", oleArgs <- oleChild , oleDomElement) )

    oleCloseAutoObject oleDomElement
}

/*! \memberof XMLDocument 
    \param xd The XMLDocument to associate with the newly created XMLTag
    \param xtChild The XMLTag object to append to the root of the document
    \brief This function will append a XMLTag that has been created by create_XMLTag(XMLDocument, string) to the document root of xd. 
    
    You need to call \b appendChild or \b insertChildBefore to make the new XMLTag show up in the document.
*/  
void appendChild (XMLDocument xd, XMLTag xtChild) { appendChildNode_ (DxlObjectOf xd, xtChild) } 

/*! \memberof XMLTag 
    \param xt The new parent of xtChild.
    \param xtChild The XMLTag object to append to the childs of xt
    \brief This function will append xtChild to the childs of xt. 
    
    xtChild needs to be created by XMLDocument::create_XMLTag(XMLDocument, string).Always make sure that xtChild and xt belong to the same 
    parent document!
*/  
void appendChild (XMLTag      xt, XMLTag xtChild) { appendChildNode_ (DxlObjectOf xt, xtChild) }

/*! \memberof XMLTag 
    \param xt an XMLTag 
    \returns A new XMLTag which contains the parent of xt
    \brief Will return the parent node of an XMLTag. \b Warning: XMLTag objects are indenpendent, different objects, even if they refer to the same node. 
    
    Keep in mind that if you call this function 5 times you will receive 5 independent XMLTag objects, that all refer to the same XML node. Destroying one
    might make the others unusuable (unless COM handles the reference counts correctly). You cannot use == to test for equality of XMLTags!    
*/  
XMLTag getParentNode (XMLTag xt) {
    OleAutoObj oleParent = null
    checkOLE (  oleGet (getOleHandle_ xt, "parentNode", oleParent)  )
    return create_XMLTag_ (getParentDocument xt, oleParent)
}

/*! \memberof XMLTag 
    \param xtNext The existing XMLTag before which the new node xtPrevious will be inserted. 
    \param xtPrevious The new node to be inserted before xtNext. 
    \brief Will insert the XMLTag xtPrevious immediately before xtNext.
    
    Use this function to insert an XML tag created by XMLDocument::create_XMLTag(XMLDocument,string) as a predecessor of (immediately before) xtNext. 
    You can use this function to insert an XMLTag everywhere in the hierarchy except as a last child. To insert an XMLTag as the last child
    use the appendChild(XMLTag, XMLTag) function.
*/  
void insertChildBefore (XMLTag xtNext, XMLTag xtPrevious ) {
    XMLTag xtParent = getParentNode xtNext
    OleAutoObj newNode
    checkOLE (  oleMethod (getOleHandle_ xtParent, "insertBefore", oleArgs <- (getOleHandle_ xtPrevious) <- (getOleHandle_ xtNext) , newNode ) ) 
    delete_XMLTag xtParent
    
    if (null newNode) { raiseError ("DOM ERROR", "Could not insert node"); }
}

/*! \memberof XMLTag 
    \param xtPrevious the XMLTag before which the new tag will be created. 
    \param tagName a name for the XML tag to be created.
    \returns A new XMLTag with the name tagName that has been inserted before the tag xtPrevious
    \brief This function will create a new tag with the name tagName and insert it before xtPrevious.

    This function combines the functions XMLDocument::create_XMLTag(XMLDocument,string) and insertChildBefore(XMLTag,XMLTag). A new tag will be created.
    and inserted before xtPrevious. 
*/  
XMLTag insertTagBefore (XMLTag xtPrevious, string tagName) {
    XMLTag xt = create_XMLTag (getParentDocument xtPrevious, tagName) 
    insertChildBefore (xtPrevious, xt) 
    return xt
}

/*! \memberof XMLTag 
    \param xtParent an XMLTag 
    \param tagName a name for the XML tag to be created.
    \returns A new XMLTag with the name tagName that has been inserted as the last child of xtParent
    \brief This function will create a new tag with the name tagName and insert it before xtPrevious.

    This function combines the functions XMLDocument::create_XMLTag(XMLDocument,string) and appendChild(XMLTag,XMLTag). A new tag will be created
    and inserted before xtPrevious. 
*/  
XMLTag appendTag (XMLTag xtParent, string tagName) {
    XMLTag xt = create_XMLTag (getParentDocument xtParent, tagName) 
    appendChild (xtParent, xt) 
    return xt
}

/*! \memberof XMLTag 
    \param xtParent an XMLDocument
    \param tagName a name for the XML tag to be created.
    \returns A new XMLTag with the name tagName that has been inserted as the last child of the root node
    \brief This function will create a new tag with the name tagName and insert it before xtPrevious.

    This function combines the functions XMLDocument::create_XMLTag(XMLDocument,string) and appendChild(XMLTag,XMLTag). A new tag will be created
    and inserted before xtPrevious. 
*/  
XMLTag appendTag (XMLDocument xtParent, string tagName) {
    XMLTag xt = create_XMLTag (xtParent, tagName) 
    appendChild (xtParent, xt) 
    return xt
}


/*! \memberof XMLTag 
    \param xt The XMLTag object which text to set. 
    \param s The new text for the XMLTag xt. 
    \brief This function will set the text of the XMLTag xt.
*/  
void   setText (XMLTag xt, string s) { checkOLE ( olePut ( getOleHandle_ xt, "text", s ) ) }

/*! \memberof XMLTag 
    \param xt The XMLTag object which text to get. 
    \returns The text for the XMLTag xt. 
    \brief This function will return the text for the XMLTag xt.
*/  
string getText (XMLTag xt          ) { string result = null; xmlCheckOLE ( oleGet ( getOleHandle_ xt, "text", result ) ); return result }

/*! \memberof XMLTag 
    \param xt The XMLTag object which value to set. 
    \param s The new value (whatever that is) for the XMLTag xt. 
    \brief This function will set the value (whatever that is) of the XMLTag xt.
*/  
void   setValue (XMLTag xt, string s) { checkOLE ( olePut ( getOleHandle_ xt, "nodeValue", s ) ) }

/*! \memberof XMLTag 
    \param xt The XMLTag object which value to get. 
    \returns The value (whatever that is) for the XMLTag xt. 
    \brief This function will return the value for the XMLTag xt.
*/  
string getValue (XMLTag xt          ) { string result = null; xmlCheckOLE ( oleGet ( getOleHandle_ xt, "nodeValue", result ) ); return result }

/*! \memberof XMLTag 
    \param xt The XMLTag object which name to get. 
    \returns The name for the XMLTag xt. 
    \brief This function will return the name for the XMLTag xt.
*/  
string getName     (XMLTag xt) { string s = null; xmlCheckOLE (oleGet (getOleHandle_ xt, "nodeName", s)); return s }


/*! \memberof XMLTag 
    \param xt The XMLTag object which attribute to set
    \param sAttrName The name of the attribute to set
    \param sAttrValue The value of the attribute to set
    \brief This function set the attribute value for the XMLTag xt.
*/  
void setAttribute (XMLTag xt, string sAttrName, string sAttrValue) {
    checkOLE ( oleMethod ( getOleHandle_ xt, "setAttribute", oleArgs <- sAttrName <- sAttrValue ) ) 
}

/*! \memberof XMLTag 
    \param xt The XMLTag object for which the attribute shall be read.
    \param sAttrName The name of the attribute that shall be read.
    \returns The value of the attribute. If the attribute does not exist, this function will return a null string. 
    \brief This function will read the attribute value of an XMLTag
*/  
string getAttribute (XMLTag xt, string sAttrName) {
    string result = ""
    string sOleResult =  oleMethod ( getOleHandle_ xt, "getAttribute", oleArgs <- sAttrName, result )
     
    if (!null sOleResult) result = null
    return result
}


/*! \memberof XMLDocument
    \param xd The XMLDocument object for which the XML shall be returned. 
    \returns The XML source of the XMLDocument
    \brief This function will return the XML for the XMLDocument object.
*/  
string getXML (XMLDocument xd) { string s = null; checkOLE (oleGet (getOleHandle_ xd, "xml", s)); return s }

/*! \memberof XMLTag
    \param xd The XMLTag object for which the XML shall be returned. 
    \returns The XML source of the XMLTag
    \brief This function will return the XML for an XMLTag and all child elements.
*/  
string getXML (XMLTag xt     ) { string s = null; checkOLE (oleGet (getOleHandle_ xt, "xml", s)); return s }


/*! \memberof XMLDocument
    \param xdTransform  The XMLDocument object for which the XML shall be returned. 
    \param xdStylesheet The XMLDocument that contains the stylesheet for the transformation
    \returns The result of the transformation
    \brief This function will apply a stylesheet to an XML document and return the transformation result
*/  
string transformDocument (XMLDocument xdTransform, XMLDocument xdStylesheet) {
    string result = ""
    checkOLE (  oleMethod (getOleHandle_ xdTransform, "transformNode", oleArgs <- getOleHandle_ xdStylesheet, result)  )
    return result
}

//! \brief XMLNodelist is a class that represents a list of XMLTags with no hierarchy
struct XMLNodelist {};

DxlObject DxlObjectOf (XMLNodelist nl) { return (addr_ nl) DxlObject }

// Properties
OleAutoObj getOleHandle_ (XMLNodelist nl)                    { DxlObject x = DxlObjectOf nl; return (x->"oleHandle") OleAutoObj }
void       setOleHandle_ (XMLNodelist nl, OleAutoObj oleObj) { DxlObject x = DxlObjectOf nl;         x->"oleHandle" = oleObj    }

/*! \memberof XMLNodelist
    \param xt The XMLNodelist object for which the parent document shall be returned. 
    \returns The parent XMLDocument object of the XMLNodelist xt
    \brief This function will return XMLDocument, which was used to create the XMLNodelist xt
*/  
XMLDocument getParentDocument   (XMLNodelist xt)                   { DxlObject x = DxlObjectOf xt; return (x->"parentDoc") XMLDocument }
void        setParentDocument_  (XMLNodelist xt, XMLDocument xd  ) { DxlObject x = DxlObjectOf xt;         x->"parentDoc" = xd         }

/*! \memberof XMLNodelist
    \param nl The XMLNodelist object for which the current node. 
    \returns The XMLTag object, that is the current node of the XMLNodelist nl
    \brief Returns the current node of the XMLNodelist nl.
    
    You can iterate through the nodes using the nextNode(XMLNodelist) function or the getItem (XMLNodelist,int) function. 
    When you call one of this functions the currentNode will change.
*/  
XMLTag     getCurrentNode  (XMLNodelist nl)            { DxlObject x = DxlObjectOf nl; return (x->"currentNode") XMLTag }

void       setCurrentNode_ (XMLNodelist nl, XMLTag xt) { DxlObject x = DxlObjectOf nl;         x->"currentNode" = xt    }


/*! \memberof XMLDocument
    \param xd The XMLDocument in which the query shall be executed
    \param expression The XPath expression query to execute on xd
    \returns An XMLNodelist object containing the nodes matching the XPATH query. 
    \brief This function will execute an XPATH query on the XMLDocument object and return the XMLTags found in a XMLNodelist.
*/  
XMLNodelist selectNodes (XMLDocument xd, string expression) {
    DxlObject x = new; XMLNodelist result = (addr_ x) XMLNodelist; 
    
    OleAutoObj oleNodeList = null
    checkOLE (  oleMethod (getOleHandle_ xd, "selectNodes", oleArgs <- expression, oleNodeList)   ) 
    if (null oleNodeList) { raiseError ("DOM ERROR", "selectNodes returned no result."); return null XMLNodelist }
    
    setOleHandle_ (result, oleNodeList) 
    setParentDocument_ (result, xd) 
    setCurrentNode_    (result, null XMLTag) 
    return result
}

/*! \memberof XMLNodelist
    \param nl XMLNodelist object to be destroyed. Will be set to NULL by the function.
    \brief This function will only destroy the XMLNodelist and its OLE handle. The contained XMLTags are deleted when the containing XMLDocument is deleted. 
    
    This function must be called for every XMLNodelist object that was created by XMLDocument::selectNodes(XMLDocument,string) or XMLTag::childNodes(XMLTag).
*/  
void delete_NodeList (XMLNodelist &nl) {
    OleAutoObj oleObj = getOleHandle_ nl; oleCloseAutoObject oleObj    

    // delete current node
    XMLTag currNode = getCurrentNode nl
    if (!null currNode) delete_XMLTag currNode
    
    DxlObject x = DxlObjectOf nl; delete x; 
    nl = null;
}


/*! \memberof XMLNodelist
    \param nl The XMLNodelist for which the number of items shall be returned. 
    \returns The number of items in the XMLNodelist nl
    \brief This function will return the number of items of an XMLNodelist. 
*/  
int count (XMLNodelist nl) {
        int count = -1
        checkOLE (  oleGet (getOleHandle_ nl, "length", count )  )
        return count 
}

/*! \memberof XMLNodelist
    \param nl The XMLNodelist for which the number of items shall be returned. 
    \param index The position of the XMLTag to receive (starting from 0)
    \returns The XMLTag object with the position index in the XMLNodelist nl
    \brief This function will return the XMLTag on position index in the XMLNodelist. The current node of nl will change. 
*/  
XMLTag getItem (XMLNodelist nl, int index) {
    OleAutoObj oleNode
    checkOLE (  oleMethod (getOleHandle_ nl, "item", oleArgs <- index, oleNode)  ) 
    if (null oleNode) { raiseError ("DOM ERROR", "Error getting item " index " on nodelist."); return null XMLTag; }
    
    return create_XMLTag_ (getParentDocument nl, oleNode)
}

/*! \memberof XMLNodelist
    \param nl The XMLNodelist for the current node shall be reset    
    \brief Resets the current node to the first node of the list
*/  
void resetIterator (XMLNodelist nl) { checkOLE ( oleMethod (getOleHandle_ nl, "reset") ) }

/*! \memberof XMLNodelist
    \param nl The XMLNodelist for the current node shall be reset    
    \returns The node after the current node
    \brief This function will return the next node in the list and \b delete the old current node object. 
    
    This function is indended for cycling through the XMLNodelist, without the need to cleanup the XMLTags 
    afterwards.
*/  
XMLTag nextNode (XMLNodelist nl)    { 
    XMLTag currNode = getCurrentNode nl
    if (!null currNode) delete_XMLTag currNode

    OleAutoObj oleNode
    checkOLE (  oleMethod (getOleHandle_ nl, "nextNode", oleArgs() , oleNode)  ) 
    if (null oleNode) { return null XMLTag; }
    return create_XMLTag_ (getParentDocument nl, oleNode)
}


/*! \memberof XMLDocument
    \param xd The XMLDocument
    \returns An XMLTag for the root element of the XML document; null if the XML document is empty
    \brief This function returns the root element object of an XMLDocument
*/
XMLTag documentElement (XMLDocument xd) {
   OleAutoObj oleNode = null
   checkOleSuccess (  oleGet (getOleHandle_ xd, "documentElement", oleNode)   )
   if (null oleNode) { return null XMLTag }
   return create_XMLTag_ (xd, oleNode)
}

/*! \memberof XMLDocument
    \param xd The XMLDocument in which the query shall be executed
    \param expression The XPath expression query to execute on xd
    \returns An XMLTag for the first match of the XPATH query. 
    \brief This function will execute an XPATH query on the XMLDocument object and return the first matching XMLTag.
*/  
XMLTag selectSingleNode (XMLDocument xd, string expression) {
   OleAutoObj oleNode = null
   checkOLE (  oleMethod (getOleHandle_ xd, "selectSingleNode", oleArgs <- expression, oleNode)   ) 
   if (null oleNode) { return null XMLTag } 
   return create_XMLTag_ (xd, oleNode)
}

/*! \memberof XMLTag
    \param xt an XMLTag
    \returns The child nodes of the XMLTag xt as a nodelist.
    \brief This function returns all child objects of an XMLTag in a XMLNodelist.
    \details The function returns also processing instructions, comments, and text nodes,
             that are child objects of the XMLTag
*/
XMLNodelist childNodes (XMLTag xt) {
    DxlObject x = new; XMLNodelist result = (addr_ x) XMLNodelist;

    OleAutoObj oleNodeList = null
    checkOleSuccess (  oleGet (getOleHandle_ xt, "childNodes", oleNodeList)   )
    if (null oleNodeList) { raiseError ("DOM ERROR", "childNodes returned no result."); return null XMLNodelist }

    setOleHandle_ (result, oleNodeList)
    setParentDocument_ (result, getParentDocument xt)
    setCurrentNode_    (result, null XMLTag)
    return result
}

/*! \memberof XMLTag
    \param xt an XMLTag
    \returns An XMLTag of the first child of the XMLTag xt 
    \brief This function returns all child objects of an XMLTag in a XMLNodelist.
*/  
XMLTag firstChild (XMLTag xt) {
   OleAutoObj oleNode = null
   checkOLE (  oleGet (getOleHandle_ xt, "firstChild", oleNode)  ) 
   if (null oleNode) { return null XMLTag } 
   return create_XMLTag_ (getParentDocument xt, oleNode)
}

/*! \memberof XMLDocument
    \return error code; 0 - if the XML document passed validation
    \param xd The XMLDocument into which the contents shall be loaded.
    \param pXMLFile The XML file from which the XML document was loaded
    \brief This function validates the XML content of an XMLDocument object.
*/
int validateXMLDocument (XMLDocument xd, string pXMLFile) {
    int iRetVal = 0
    bool result = false

    Buffer bParseMsg = null
    Buffer bError = null

    if (!null pXMLFile)
    {
        if (save_XMLDocument_(xd, pXMLFile) == 0)
        {
            if (checkOleSuccess(  oleMethod (getOleHandle_ xd, "load", oleArgs <- pXMLFile, result)  ))
            {
                bParseMsg = getParseErrorMsg_ xd
                if (!null bParseMsg)
                {
                    bError = create
                    bError += "XMLVALERR for "
                    bError += "<"
                    bError += pXMLFile
                    bError += ">:\n"
                    bError += "Validation of XML file failed with following error(s):\n"
                    /*
                     * We have encountered errors during parsing of the XML file.
                     * Reload and perform a full validation afterwards, if possible:
                     */
                    checkOleSuccess ( olePut (getOleHandle_ xd, "validateOnParse", false) )
                    checkOleSuccess(  oleMethod (getOleHandle_ xd, "load", oleArgs <- pXMLFile)  )
                }
                else
                {
                    return iRetVal
                }
            }
        }
    }

    bParseMsg = getValidationErrorMsg_(xd, bParseMsg)

    if (null bError && length(bParseMsg) > 0)
    {
        bError = create
        bError += "Validation of XML document failed with following error(s):\n"
    }

    if (!null bError)
    {
        combine(bError, bParseMsg, 0)
        print tempStringOf bError
        iRetVal = XMLVALERR
    }

    delete bParseMsg
    delete bError

    return iRetVal
}

/*! \memberof XMLDocument
    \return error code; 0 - if the XML document passed validation
    \param xd The XMLDocument into which the contents shall be loaded.
    \brief This function validates the XML content of an XMLDocument object.
*/
int validateXMLDocument (XMLDocument xd) {
    string pLocation = ""

    checkOleSuccess (  oleGet (getOleHandle_ xd, "url", pLocation)   )

    return validateXMLDocument (xd, pLocation)
}

