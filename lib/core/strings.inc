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

/*! \file strings.inc 
    \brief Basic string manipulation functions
 
This file will contain functions for string manipulation, like splitting, joining, sorting and many more. Check 
in later to see the progress.
*/


/*! \return the contents of s where leading whitespaces (space,\\n,\\t,\\r) are removed from the beginning and the end
    \param s the string where leading and trailing whitespaces shall be removed
    \brief efficient removal of whitespaces at beginning and end
    This function is designed so that no unnecessary intermediate strings are created. That is why it runs on 'char' level.
*/  
string trimWhitespace(string s) {

   int first = 0
   int last = ( length s ) - 1
   while ( last > 0 && isspace( s[ last ] ) ) last--
   while ( isspace( s[ first ] ) && first < last ) first++
   if ( first == last && isspace s[first] ) return ""
   return s[ first:last ]
}

/*! \example example_trimWhitespace.dxl 
    This is an example on how to use the trimWhitespace(string) function. */

/*! \return Returns sSource, where the searchstring \em sSearch is replaced by sReplace 
    \param sSource the source string in which \em sSearch shall be replaced
    \param sSearch the search string that shall be replaced
    \param sReplace the string by which \em sSearch shall be replaced
    \brief Replaces a part of the string by another string
*/

Buffer gbuf_replace_search = create(10000);
Buffer gbuf_replace_source = create(10000);
Buffer gbuf_replace_result = create(10000);

string replace (string sSource, string sSearch, string sReplace) {
    gbuf_replace_source = sSource; 
    gbuf_replace_search = sSearch;
    int iLen = length gbuf_replace_source
    if (iLen == 0) return ""
    
    int iLenSearch = length(gbuf_replace_search);
    if (iLenSearch == 0) { 
        raiseError ("Parameter error", "in strings.inc/replace: search string must not be empty")
        return "" 
    }
    
    // read the first char for latter comparison -> speed optimization
    char firstChar = gbuf_replace_search[0]

    int pos = 0, d1,d2;
    int i
    
    gbuf_replace_result = "";    
    while (pos < iLen) { 
        char ch = gbuf_replace_source[pos]; 
        bool found = true
        if (ch != firstChar) { pos ++; gbuf_replace_result += ch; continue }
        for (i = 1; i < iLenSearch; i++) {
           if (gbuf_replace_source[pos+i] != gbuf_replace_search[i]) { found = false; break }
        }
        if (!found) { pos++; gbuf_replace_result += ch; continue}
        gbuf_replace_result += sReplace
        pos += iLenSearch
    }
    
    string result = stringOf gbuf_replace_result
    return result
}

/**
 *  bool string_replace(Regexp pPattern, string pReplaceWith, int iStart, int iEnd, Buffer &p_buff)
 *  @brief Replaces a given pattern within text buffer
 *  @param[in] pPattern Regular expression with the pattern to search for
 *  @param[in] pReplaceWith The replace string.
 *  @param[in] iStart (optional) Position in the buffer, where to start the replacement (iStart >= 0)
 *  @param[in] iEnd (optional)   Position in the buffer, where to end the replacement
 *  @param[in,out] p_buff Text buffer
 *  @return <b>bool</b> False, if the pattern has not been found in the buffer
 *
 *  @note This function handles iStart and iEnd in the same way as DOORS DXL function
 *        search():
 *        iStart < 0:              not allowed; no replacement
 *        iEnd < 0:                ignore parameter iEnd
 *        iEnd < iStart:           ignore parameter iEnd
 *        iStart >= length buffer: no replacement
 *
 *  @par Example:
 *  @code
 *  Regexp regPattern = regexp2 "Test
 *  Buffer b = create
 *  b = "This is a Test string."
 *  // ...
 *  if (string_replace(regPattern, "new test", b)
 *  {
 *      print "Result: "
 *      print tempStringOf(b)
 *      print "\n"
 *  }
 *  // ...
 *  delete regPattern
 *  @endcode
 *
 *  Original string (original): This is a Test string.<br/>
 *  Pattern (pattern): Test<br/>
 *  Replace with (replaceWith): new Test<br/>
 *  <br/>
 *  Result: This is a new test string.
 */
bool string_replace(Regexp pPattern, string pReplaceWith, int iStart, int iEnd, Buffer &p_buff)
{
    bool doneReplace = false
    Buffer tmp_buf = create()
    int last_index = iStart
    int current_index = 0
    int offset = 0

    // Functon search() is pretty tolerant regarding start/end values
    // But as we repeatedly use it, we need to fix some strange settings,
    // that search() usually accepts
    if (last_index > iEnd)
    {
        // search() ignores iEnd if iEnd < iStart
        iEnd = -1
    }
    if (last_index > 0)
    {
         combine(tmp_buf, p_buff, 0, last_index - 1)
    }
    while (search(pPattern, p_buff, last_index, iEnd)) {
        offset = end 0 + 1
        current_index = last_index + start 0 - 1
        if (current_index >= last_index)
        {
            combine(tmp_buf, p_buff, last_index, current_index)
        }
        last_index = last_index + offset
        tmp_buf += pReplaceWith
        doneReplace = true
        if (iEnd >= 0 && last_index > iEnd)
        {
            // We reached the end of the replace area, but can't rely
            // on function search(), as it would now start to ignore iEnd
            break;
        }
    }
    if (last_index < 0)
    {
        combine(tmp_buf, p_buff, 0)
    }
    else
    {
        combine(tmp_buf, p_buff, last_index)
    }
    delete(p_buff)
    p_buff = tmp_buf
    return doneReplace
}

/**
 *  @overload bool string_replace(Regexp pPattern, string pReplaceWith, int iStart, Buffer &p_buff)
 *  @brief Replaces a given pattern within text buffer
 *
 *  @note If also an end postion for the replacement is required, use the function
 *        overload above.
 *
 */
bool string_replace(Regexp pPattern, string pReplaceWith, int iStart, Buffer &p_buff)
{
    int iEnd = -1 // search and replace from iStart to the end of the buffer
    return string_replace(pPattern, pReplaceWith, iStart, iEnd, p_buff)
}

/**
 *  @overload bool string_replace(Regexp pPattern, string pReplaceWith, Buffer &p_buff)
 *  @brief Replaces a given pattern within a text buffer
 *
 *  @note If also a start and/or end postion for the replacement is required, use the function
 *        overload above.
 *
 */
bool string_replace(Regexp pPattern, string pReplaceWith, Buffer &p_buff)
{
    int iStart = 0  // Search and replace from the beginning of the buffer
    int iEnd = -1   // until its end
    return string_replace(pPattern, pReplaceWith, iStart, iEnd, p_buff)
}
