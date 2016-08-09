#tag Module
Protected Module VB
	#tag Method, Flags = &h1
		Protected Sub AppActivate(title As String, unused As Boolean = False)
		  // We want to activate an "application" based on the
		  // window title, or partial window title.  From the VB6 docs:
		  //
		  // In determining which application to activate, title is compared to the title 
		  // string of each running application. If there is no exact match, any application 
		  // whose title string begins with title is activated. If there is more than one 
		  // instance of the application named by title, one instance is arbitrarily activated.
		  //
		  // The way I interpret this is that we should loop over all the processes
		  // and look at the name.  If the name is an exact match, then we use
		  // that one.  If no matches are found, we go back and look at the 
		  // beginning of the each process name to see if we can find a match.  If
		  // we still can't find one, then I think we're supposed search window titles
		  // in the same fashion.  It's ambiguous though.  The parameter specifier
		  // from the same docs says: "...the title in the title bar of the application window you
		  // want to activate..."  So which is it -- process name or window title?
		  //
		  // Turns out that the answer is "exact window title", as learned from
		  // http://support.microsoft.com/kb/q147659/
		  //
		  // "The Visual Basic AppActivate command can only activate a window if 
		  // you know the exact window title. Similarly, the Windows FindWindow 
		  // function can only find a window handle if you know the exact window title. "
		  
		  #If TargetWin32
		    Soft Declare Function FindWindowA Lib "User32" (className As Integer, title As CString) As Integer
		    Soft Declare Function FindWindowW Lib "User32" (className As Integer, title As WString) As Integer
		    
		    // Find the window via an exact match using FindWindow
		    Dim handle As Integer
		    If System.IsFunctionAvailable("FindWindowW", "User32") Then
		      handle = FindWindowW(0, title)
		    Else
		      handle = FindWindowA(0, title)
		    End If
		    
		    // If we found a handle, then we want to bring it to the front
		    If handle <> 0 Then
		      Declare Sub SetForegroundWindow Lib "User32" (hwnd As Integer)
		      SetForegroundWindow(handle)
		    End If
		  #EndIf
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function ASCIIToScanKey(char As String) As Integer
		  #If TargetWin32
		    Declare Function VkKeyScanA Lib "User32" (ch As Integer) As Int16
		    
		    Dim theAscVal As Integer = Asc(char)
		    Return VkKeyScanA(theAscVal)
		  #EndIf
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function ATn(d As Double) As Double
		  Return ATan(d)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ChDir(path As Text)
		  #If TargetWin32
		    Soft Declare Function SetCurrentDirectoryW Lib "Kernel32" (dir As WString) As Boolean
		    
		    Call SetCurrentDirectoryW(path)
		  #Endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub ChDrive(letter As Text)
		  // We only want to look at the first character, so let's
		  // ensure that we have only one.
		  #If TargetWindows
		    letter = letter.Left(1)
		    
		    // If the letter is empty, we can bail out
		    If letter = "" Then Return
		    
		    // Now we want to change the drive.  We can
		    // do this with ChDir.
		    ChDir(letter + ":\")
		  #Endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Command() As Text
		  // This should return just the command line part after the file
		  // name.  It returns it without parsing it.  In Xojo, the
		  // command line comes with the filename, so we need to parse
		  // that out.
		  
		  // Get the entire command line
		  Dim commandLine As String = System.CommandLine
		  
		  // Let the helper function deal with the hard stuff
		  Dim temp As String
		  Return GetParams(commandLine, temp).ToText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function CurDir() As Text
		  // The user just wants the path to the current directory
		  Return GetCurrentDirectory.NativePath.ToText
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Date() As Text
		  // Get the current long date as text
		  Return Xojo.Core.Date.Now.ToText(Xojo.Core.Locale.Current, _
		  Xojo.Core.Date.FormatStyles.Long, Xojo.Core.Date.FormatStyles.None)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Day() As Integer
		  // Returns the day of the month for the current date
		  Return Xojo.Core.Date.Now.Day
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub DeleteSetting(appName As String, section As String, key As String = "")
		  #If TargetWindows
		    // First, we want to get a registry key that points
		    // to the default location of all VB settings.
		    Dim base As New RegistryItem(kSettingsLocation)
		    
		    // Now we want to delve into the appName folder
		    base = base.Child(appName)
		    
		    // If we don't have a key name, then we want to
		    // delete the entire section.  Otherwise, we want
		    // to delve into the section and delete the
		    // key specified.
		    If key = "" Then
		      base.Delete(section)
		    Else
		      // Dive into the section
		      base = base.Child(section)
		      // And delete the key
		      base.Delete(key)
		    End If
		    
		    
		    Exception err As RegistryAccessErrorException
		      // Something bad happened, so let's just bail out
		      Return
		  #Endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub FileCopy(source As Text, dest As Text)
		  Dim sourceItem As New Xojo.IO.FolderItem(source)
		  Dim destItem As New Xojo.IO.FolderItem(dest)
		  
		  // Copy the source to the dest
		  Try
		    sourceItem.CopyTo(destItem)
		  Catch e As IOException
		    Return
		  End Try
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub FillKeyMap(ByRef map As Dictionary)
		  map.Value("BACKSPACE") = &h8
		  map.Value("BS") = &h8
		  map.Value("BKSP") = &h8
		  map.Value("BREAK") = &h3
		  map.Value("CAPSLOCK") = &h14
		  map.Value("DELETE") = &h2E
		  map.Value("DEL") = &h2E
		  map.Value("DOWN") = &h28
		  map.Value("END") = &h23
		  map.Value("ENTER") = &h0D
		  map.Value("ESC") = &h1B
		  map.Value("HELP") = &h2F
		  map.Value("HOME") = &h24
		  map.Value("INSERT") = &h2D
		  map.Value("INS") = &h2D
		  map.Value("LEFT") = &h25
		  map.Value("NUMLOCK") = &h90
		  map.Value("PGDN") = &h22
		  map.Value("PGUP") = &h21
		  map.Value("PRTSC") = &h2C
		  map.Value("RIGHT") = &h27
		  map.Value("SCROLLLOCK") = &h91
		  map.Value("TAB") = &h09
		  map.Value("UP") = &h26
		  
		  map.Value("+") = ASCIIToScanKey("+")
		  map.Value("^") = ASCIIToScanKey("^")
		  map.Value("%") = ASCIIToScanKey("%")
		  map.Value("~") = ASCIIToScanKey("~")
		  map.Value("(") = ASCIIToScanKey("(")
		  map.Value(")") = ASCIIToScanKey(")")
		  map.Value("{") = ASCIIToScanKey("{")
		  map.Value("}") = ASCIIToScanKey("}")
		  map.Value("[") = ASCIIToScanKey("[")
		  map.Value("]") = ASCIIToScanKey("]")
		  
		  For i As Integer = 1 To 16
		    map.Value("F" + Str(i)) = &h70 + (i - 1)
		  Next i
		  
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function FillString(char As String, numChars As Integer) As String
		  #If TargetWin32
		    Declare Sub memset Lib "msvcrt" (dest As Ptr, char As Integer, count As Integer)
		    
		    Dim mb As New MemoryBlock(LenB( char ) * numChars)
		    memset(mb, AscB(char), numChars)
		    
		    Return DefineEncoding(mb, Encoding(char))
		  #EndIf
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Filter(source() As String, match As String, include As Boolean = True, compare As Integer = 1) As String()
		  // We want to filter the entries from source which match the match 
		  // string.  The include flag says whether we want to return entries
		  // which do match, or which don't match.  The compare flag says
		  // what type of comparison to use.
		  
		  Dim ret(-1) As String
		  For Each s As String In source
		    Dim add As Boolean
		    Select Case compare
		    Case 0, 1  // Binary or text comparison
		      If StrComp(s, match, compare) = 0 Then add = True
		    Else
		      If s = match Then add = True
		    End Select
		    
		    // If we're doing exclusion, then add means
		    // we don't want to add it
		    If Not include And add Then add = False
		    
		    // If we want to add it, then do it
		    If add Then ret.Append(s)
		  Next s
		  
		  // Return the results
		  Return ret
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Fix(num As Double) As Integer
		  // This function returns the integer portion of the number passed.
		  // If the number is negative, Fix returns the first negative integer
		  // greater than or equal to the number.  For example, Fix converts -8.4
		  // to -8.  If you want -9, then you should be using Int instead.
		  
		  Return Sgn(num) * Int(Abs(num))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function FV(rate As Double, nper As Integer, pmt As Double, pv As Double = 0, type As Integer = 0) As Double
		  // These equations come from gnucash
		  // http://www.gnucash.org/docs/v1.8/C/gnucash-guide/loans_calcs1.html
		  Dim a As Double = (1 + rate) ^ nper - 1
		  Dim b As Double = (1 + rate * type) / rate
		  Dim c As Double = pmt * b
		  
		  Return -(pv + a * (pv + c))
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetAllSettings(appName As String, section As String) As Dictionary
		  #If TargetWindows
		    // First, we want to get to the default location for all VB apps
		    Dim base As New RegistryItem(kSettingsLocation)
		    
		    // Then we want to dive into the app and section folders
		    base = base.Child(appName).Child(section)
		    
		    // Loop over all the children and return their values
		    Dim i, count As Integer
		    Dim ret As New Dictionary
		    
		    // How many keys do we have?
		    count = base.KeyCount
		    
		    For i = 0 To count - 1
		      // Grab the key and value and add it to the dictionary
		      ret.Value(base.Name(i)) = base.Value(i)
		    Next i
		    
		    // Return our list
		    Return ret
		    
		    Exception err As RegistryAccessErrorException
		      // Just bail out
		      Return Nil
		  #Endif
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetCurrentDirectory() As FolderItem
		  #If TargetWin32
		    Soft Declare Sub GetCurrentDirectoryA Lib "Kernel32" (Len As Integer, buf As Ptr)
		    Soft Declare Sub GetCurrentDirectoryW Lib "Kernel32" (Len As Integer, buf As Ptr)
		    
		    Dim path As String
		    Dim buf As New MemoryBlock(520)
		    If System.IsFunctionAvailable("GetCurrentDirectoryW", "Kernel32") Then
		      GetCurrentDirectoryW(buf.Size, buf)
		      path = buf.WString(0)
		    Else
		      GetCurrentDirectoryA(buf.Size, buf)
		      path = buf.CString(0)
		    End If
		    
		    Return New FolderItem(path, FolderItem.PathTypeAbsolute)
		    
		  #EndIf
		  
		  Exception err As UnsupportedFormatException
		    Return Nil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Function GetParams(commandLine As String, ByRef file As String) As String
		  // Now parse the command line so that we can get just what
		  // is after the application name.  We have to do this one
		  // character at a time, unfortunately
		  Dim length As Integer = Len(commandLine)
		  Dim ignoreSpaces As Boolean = False
		  
		  For curPos As Integer = 1 To length
		    Dim char As String = Mid(commandLine, curPos, 1)
		    
		    If char = """" Then
		      // We found a quote, so we can ignore any spaces.  If
		      // this is the second quote, then we can pay attention
		      // to spaces again.
		      ignoreSpaces = Not ignoreSpaces
		    ElseIf Not ignoreSpaces And (char = " " Or char = Chr(9)) Then
		      // We have a space.  If we're ignoring spaces, then
		      // it doesn't matter.  But if we're not, then we've found
		      // the end of the application name
		      file = Trim(Left(commandLine, curPos))
		      Return Trim(Mid(commandLine, curPos))
		    End If
		  Next curPos
		  
		  file = commandLine
		  Return ""
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function GetSetting(appName As String, section As String, key As String, default As Variant = "") As Variant
		  #If TargetWindows
		    // First, we want to get to the default location for all VB apps
		    Dim base As New RegistryItem(kSettingsLocation)
		    
		    // Then we want to dive into the app and section folders
		    base = base.Child(appName).Child(section)
		    
		    // Now we want to get the value from the key
		    Return base.Value(key)
		    
		    Exception err As RegistryAccessErrorException
		      // Just bail out
		      Return default
		      
		  #Endif
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Hour() As Integer
		  // Return the hour for the current time
		  Return Xojo.Core.Date.Now.Hour
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function InStrRev(source As String, substr As String, startPos As Integer = 1, compare As Integer = 1) As Integer
		  If source = "" Then Return 0
		  If substr.Len = 0 Then Return startPos
		  // Similar to InStr, but searches backwards from the given position
		  // (or if startPos = -1, then from the end of the string).
		  // If substr can't be found, returns 0.
		  
		  Dim srcLen As Integer
		  If compare = 0 Then
		    srcLen = source.LenB
		  Else
		    srcLen = source.Len
		  End If
		  
		  If startPos > srcLen Then Return 0
		  
		  // Here's an easy way...
		  // There may be a faster implementation, but then again, there may not -- it probably
		  // depends on what you're trying to do.
		  Dim reversedSource As String = StrReverse(source)
		  Dim reversedSubstr As String = StrReverse(substr)
		  Dim reversedPos As Integer
		  If compare = 0 Then
		    reversedPos = InStrB(startPos, reversedSource, reversedSubstr)
		  Else
		    reversedPos = InStr(startPos, reversedSource, reversedSubstr)
		  End If
		  If reversedPos < 1 Then Return 0
		  
		  If compare = 0 Then
		    Return srcLen - reversedPos - substr.LenB + 2
		  Else
		    Return srcLen - reversedPos - substr.Len + 2
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Int(num As Double) As Integer
		  // This function returns the integer portion of the number passed.
		  // If the number is negative, Int returns the first negative integer 
		  // less than or equal to the number.  For example, Int converts -8.4 
		  // to -9.  If you want -8, then you should be using Fix instead.
		  
		  Dim i As Integer = num
		  If num > 0 Then
		    Return i
		  Else
		    Return i - 1
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IPmt(rate As Double, per As Integer, nper As Integer, pv As Double, fv As Double = 0, type As Integer = 0) As Double
		  // IPmt is the principle for the previous month times the interest rate
		  // http://www.gnome.org/projects/gnumeric/doc/gnumeric-IPMT.shtml
		  
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function IsObject(v As Variant) As Boolean
		  // If the variant holds an object, this is true.  Also, if
		  // it holds nil, then it's true as well.
		  Return v.Type = Variant.TypeObject Or v.Type = Variant.TypeNil
		End Function
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub KeyDown(virtualKeyCode As Integer, extendedKey As Boolean = False)
		  #If TargetWin32
		    Declare Sub keybd_event Lib "User32" (keyCode As Integer, scanCode As Integer, _
		    flags As Integer, extraData As Integer)
		    
		    Dim flags As Integer
		    Const KEYEVENTF_EXTENDEDKEY = &h1
		    If extendedKey Then
		      flags = KEYEVENTF_EXTENDEDKEY
		    End If
		    
		    ' Press the key
		    keybd_event(virtualKeyCode, 0, flags, 0)
		  #EndIf
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h21
		Private Sub KeyUp(virtualKeyCode As Integer, extendedKey As Boolean = False)
		  #If TargetWin32
		    Declare Sub keybd_event Lib "User32" (keyCode As Integer, scanCode As Integer, _
		    flags As Integer, extraData As Integer)
		    
		    Dim flags As Integer
		    Const KEYEVENTF_EXTENDEDKEY = &h1
		    If extendedKey Then
		      flags = KEYEVENTF_EXTENDEDKEY
		    End If
		    
		    Const KEYEVENTF_KEYUP = &h2
		    flags = BitwiseOr(flags, KEYEVENTF_KEYUP)
		    keybd_event(virtualKeyCode, 0, flags, 0)
		  #EndIf
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Kill(path As String)
		  // This deletes files from the disk.  Also, 
		  // it supports wildcard characters such as * (for multiple
		  // characters) and ? (for single characters) as a way to 
		  // specify multiple files.
		  
		  // We need to use FindFirstFile as a way to find all the
		  // files that we want to delete.  We will come up with a
		  // list of FolderItems, and then we can just use 
		  // FolderItem.Delete on them.
		  
		  // Make sure the path points to our current directory as well
		  Dim curDir As String = GetCurrentDirectory.NativePath
		  path = curDir + path
		  
		  Dim toBeDeleted() As FolderItem
		  #If TargetWin32
		    Soft Declare Function FindFirstFileA Lib "Kernel32" (name As CString, data As Ptr) As Integer
		    Soft Declare Function FindFirstFileW Lib "Kernel32" (name As WString, data As Ptr) As Integer
		    Soft Declare Function FindNextFileA Lib "Kernel32" (handle As Integer, data As Ptr) As Boolean
		    Soft Declare Function FindNextFileW Lib "Kernel32" (handle As Integer, data As Ptr) As Boolean
		    Declare Sub FindClose Lib "Kernel32" (handle As Integer)
		    
		    // Check to see whether we're doing unicode processing or not
		    Dim isUnicode As Boolean = False
		    If System.IsFunctionAvailable("FindNextFileW", "Kernel32") Then isUnicode = True
		    
		    // Get the search handle
		    Dim searchHandle As Integer
		    Dim searchData As New MemoryBlock(44 + 520 + 28)
		    If isUnicode Then
		      searchHandle = FindFirstFileW(path, searchData)
		    Else
		      searchHandle = FindFirstFileA(path, searchData)
		    End If
		    
		    // If the search handle is 0, then we know that something's wrong and
		    // we can bail out
		    If searchHandle = 0 Then Return
		    
		    // Loop over all the files and add them to our kill list
		    Dim done As Boolean
		    Do
		      // Add the file to our delete list
		      Try
		        If isUnicode Then
		          toBeDeleted.Append(New FolderItem(curDir + searchData.WString(44), FolderItem.PathTypeNative))
		        Else
		          toBeDeleted.Append(New FolderItem(curDir + searchData.CString(44), FolderItem.PathTypeNative))
		        End If
		      Catch err As UnsupportedFormatException
		        // We had an error, but I think we should keep trying.
		      End Try
		      
		      // Find the next file in our list
		      If isUnicode Then
		        done = Not FindNextFileW(searchHandle, searchData)
		      Else
		        done = Not FindNextFileA(searchHandle, searchData)
		      End If
		    Loop Until done
		    
		    // Close the search handle
		    FindClose(searchHandle)
		  #EndIf
		  
		  // Now we can loop over all the files to be deleted
		  // and delete them
		  For Each item As FolderItem In toBeDeleted
		    item.Delete
		  Next item
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function LCase(s As String) As String
		  Return s.Lowercase
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Like(toSearch As String, matchingPattern As String) As Boolean
		  Static re As RegEx
		  
		  If re = Nil Then re = New RegEx
		  
		  // convert Like syntax to RegEx syntax
		  matchingPattern = matchingPattern.ReplaceAll(".", "\.")
		  matchingPattern = matchingPattern.ReplaceAll("*", ".*")
		  matchingPattern = matchingPattern.ReplaceAll("#","\d")
		  matchingPattern = matchingPattern.ReplaceAll("[!", "[^")
		  
		  // special replace for "[x]" syntax in Like
		  re.SearchPattern = "\[(.)\]" // match 1 char in brackets
		  re.ReplacementPattern = "\\\1"
		  re.Options.ReplaceAllMatches = True
		  matchingPattern = re.Replace(matchingPattern)
		  
		  // special replace for "?"
		  re.SearchPattern = "(?<!\\)\?"
		  re.ReplacementPattern = "."
		  matchingPattern = re.Replace(matchingPattern)
		  
		  // now set up RegEx
		  re.SearchPattern = "^" + matchingPattern + "$"
		  
		  // and see if it matches toSearch
		  If Nil = re.Search(toSearch) Then
		    // no match found?
		    Return False
		  Else
		    // it did match
		    Return True
		  End If
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub LSet(ByRef dest As String, assigns source As String)
		  // We want to take the source string and left-align it in the
		  // destination string.  What this essentially does is puts the 
		  // source in the left-hand part of dest, and fills the rest of
		  // dest with spaces.  So:
		  // 
		  // Dim MyString as String = "0123456789"
		  // Lset MyString = "<-Left"   
		  //
		  // Means that MyString contains "<-Left    ".
		  
		  // First, calculate the end length of the destination
		  Dim destLen As Integer = Len(dest)
		  Dim sourceLen As Integer = Len(source)
		  
		  // If the source string is greater than the
		  // destination, we want to trim the source
		  // string and just be done
		  If sourceLen >= destLen Then
		    dest = Left(source, destLen)
		    Return
		  End If
		  
		  // Otherwise, we're stuck doing it the "hard" way.
		  // First, assign the source (this would make it left-aligned).
		  dest = source
		  
		  // Then calculate how many spaces we need to 
		  // add to fill the rest of the length
		  Dim numSpaces As Integer = destLen - sourceLen
		  
		  // Now add the spaces
		  dest = dest + Space(numSpaces)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Mid(ByRef txt As String, startPos As Integer, length As Integer = -1, assigns subStr As String)
		  // Assign the replacement to the original data
		  txt = Left(txt, startPos) + Left(subStr, length) + _
		  Mid(txt, startPos + length + 1)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Minute() As Integer
		  // Return the current minute
		  Return Xojo.Core.Date.Now.Minute
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub MkDir(name As String)
		  // First, get the current directory
		  Dim curDir As FolderItem = GetCurrentDirectory
		  If curDir = Nil Then Return
		  
		  // Now, we want to make a new directory as a
		  // child of the current one
		  Dim newDir As FolderItem = curDir.Child(name)
		  newDir.CreateAsFolder
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Month() As Integer
		  // Return the current month
		  Return Xojo.Core.Date.Now.Month
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Name(oldPathName As String, newPathName As String)
		  // This moves and/or renames a file.  Sound familiar?
		  // Check to see whether the old path or the new path
		  // really are paths.  If they're not, we need to use the
		  // current directory.
		  
		  Dim oldPathIsAbsolute, newPathIsAbsolute As Boolean
		  If Mid(oldPathName, 2, 2) = ":\" Or Left(oldPathName, 2) = "//" Then
		    oldPathIsAbsolute = True
		  End If
		  
		  If Mid(newPathName, 2, 2) = ":\" Or Left(newPathName, 2) = "//" Then
		    newPathIsAbsolute = True
		  End If
		  
		  // Now we can get folder items for both the new and the
		  // old path.
		  Dim oldPath As FolderItem
		  If oldPathIsAbsolute Then
		    oldPath = New FolderItem(oldPathName, FolderItem.PathTypeAbsolute)
		  Else
		    oldPath = GetCurrentDirectory.Child(oldPathName)
		  End If
		  
		  Dim newPath As FolderItem
		  If newPathIsAbsolute Then
		    newPath = New FolderItem(newPathName, FolderItem.PathTypeAbsolute)
		  Else
		    newPath = GetCurrentDirectory.Child(newPathName)
		  End If
		  
		  // Now we can do a move operation.  This will also do a rename if
		  // the oldPath and the newPath reside in the same directory
		  oldPath.MoveFileTo(newPath)
		  
		  Exception err As UnsupportedFormatException
		    Return
		  Exception err As NilObjectException
		    Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Now() As Xojo.Core.Date
		  Return Xojo.Core.Date.Now
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Pmt(rate As Double, nper As Integer, pv As Double, fv As Double = 0, type As Integer = 0) As Double
		  // These equations come from gnucash
		  // http://www.gnucash.org/docs/v1.8/C/gnucash-guide/loans_calcs1.html
		  Dim a As Double = (1 + rate) ^ nper - 1
		  Dim b As Double = (1 + rate * type) / rate
		  
		  Return -(fv + pv * (a + 1)) / (a * b)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PPmt(rate As Double, per As Integer, nper As Integer, pv As Double, fv As Double = 0, type As Integer = 0) As Double
		  // PPmt is just the Pmt - IPmt, according to 
		  // http://www.gnome.org/projects/gnumeric/doc/gnumeric-PPMT.shtml
		  Return Pmt(rate, nper, pv, fv, type) - IPmt(rate, per, nper, pv, fv, type)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function PV(rate As Double, nper As Integer, pmt As Double, fv As Double = 0, type As Integer = 0) As Double
		  // These equations come from gnucash
		  // http://www.gnucash.org/docs/v1.8/C/gnucash-guide/loans_calcs1.html
		  Dim a As Double = (1 + rate) ^ nper - 1
		  Dim b As Double = (1 + rate * type) / rate
		  Dim c As Double = pmt * b
		  
		  Return -(fv + a * c) / (a + 1)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function QBColor(index As Integer) As Integer
		  Select Case index
		  Case 0  // Black
		    Return &h000000
		  Case 1  // Blue
		    Return &h800000
		  Case 2  // Green
		    Return &h8000
		  Case 3  // Cyan
		    Return &h808000
		  Case 4  // Red
		    Return &h80
		  Case 5  // Magenta
		    Return &h800080
		  Case 6  // Yellow
		    Return &h8080
		  Case 7  // White
		    Return &hC0C0C0
		  Case 8  // Gray
		    Return &h808080
		  Case 9  // Light blue
		    Return &hFF0000
		  Case 10  // Light green
		    Return &hFF00
		  Case 11  // Light cyan
		    Return &hFFFF00
		  Case 12  // Light red
		    Return &hFF
		  Case 13  // Light magenta
		    Return &hFF00FF
		  Case 14  // Light yellow
		    Return &hFFFF
		  Case 15  // Bright white
		    Return &hFFFFFF
		  Else
		    Return 0
		  End Select
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Randomize(seed As Integer = -1)
		  If seed <> -1 Then
		    mRnd.Seed = seed
		  Else
		    mRnd.Seed = mRnd.Number * &hFFFFFFFF
		  End If
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Replace(source As String, find As String, rep As String, start As Integer = 1, count As Integer = -1, compare As Integer = 1) As String
		  // We want to replace the search string a certain number of times
		  // in the source string.  This is different than our Replace function which
		  // only replaces the first instance and ReplaceAll, which replaces all
		  // instances.
		  
		  // Do our santity checks
		  If source = "" Then Return ""
		  If find = "" Then Return source
		  If rep = "" Then Return source
		  If count = 0 Then Return source
		  
		  // If the user wants to start farther up the string than at
		  // the first character, we need to do some wiggling since
		  // REALbasic doesn't let you do specify a start position for
		  // the source string in Replace
		  Dim searchStr As String = Mid(source, start)
		  'Dim curPos As Integer = 1
		  
		  If count = -1 Then
		    // We just want to do a replace all
		    If compare = 0 Then
		      searchStr = ReplaceAllB(searchStr, find, rep)
		    Else
		      searchStr = ReplaceAll(searchStr, find, rep)
		    End If
		  Else
		    // Now we want to do the replaces over and over again.
		    While count > 0
		      If compare = 0 Then
		        searchStr = ReplaceB(searchStr, find, rep)
		      Else
		        searchStr = Replace(searchStr, find, rep)
		      End If
		      
		      // We have one less replace to do
		      count = count - 1
		    Wend
		  End If
		  
		  // Now we're set.  The only thing we might have to do
		  // is reconstitute the original part of the search string if
		  // start is greater than 1.
		  If start > 1 Then
		    Return Left(source, start - 1) + searchStr
		  Else
		    Return searchStr
		  End If
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RmDir(path as String)
		  // Check to see if the path is an absolute path, or
		  // just a local one
		  Dim itemToDelete As FolderItem
		  If Mid(path, 2, 2) = ":\" Or Left(path, 2) = "//" Then
		    itemToDelete = New FolderItem(path, FolderItem.PathTypeNative)
		  Else
		    itemToDelete = GetCurrentDirectory.Child(path)
		  End If
		  
		  // Then delete the item
		  itemToDelete.Delete
		  
		  Exception err As UnsupportedFormatException
		    Return
		  Exception err As NilObjectException
		    Return
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Rnd() As Double
		  Return mRnd.Number
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub RSet(ByRef dest As String, assigns source As String)
		  // This is the sibling to LSet.
		  //
		  // Dim MyString as String = "0123456789"
		  // Rset(MyString) = "Right->"
		  //
		  // MyString contains "   Right->".
		  
		  // First, calculate the end length of the destination
		  Dim destLen As Integer = Len(dest)
		  Dim sourceLen As Integer = Len(source)
		  
		  // If the source string is greater than the
		  // destination, we want to trim the source
		  // string and just be done
		  If sourceLen >= destLen Then
		    dest = Right(source, destLen)
		    Return
		  End If
		  
		  // Otherwise, we're stuck doing it the "hard" way.
		  
		  // First, calculate how many spaces we need to
		  // add to fill the rest of the length
		  Dim numSpaces As Integer = destLen - sourceLen
		  
		  // Now add the spaces
		  dest = Space(numSpaces)
		  
		  // Then, assign the source (this would make it right-aligned).
		  dest = dest + source
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SavePicture(p As Picture, name As String)
		  // We want to save the given picture in a file, which
		  // Xojo pretty much already handles for you.
		  
		  // Check to see if the path is an absolute path, or
		  // just a local one
		  Dim fileToSave As FolderItem
		  fileToSave = GetCurrentDirectory.Child(name)
		  
		  // Then save the picture out
		  fileToSave.SaveAsPicture(p)
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SaveSetting(appName As String, section As String, key As String, setting As Variant)
		  #If TargetWindows
		    // First, we want to get to the default location for all VB apps
		    Dim base As New RegistryItem(kSettingsLocation)
		    
		    // Then we want to dive into the app and section folders
		    base = base.Child(appName).Child(section)
		    
		    // Now we want to save the key and value
		    base.Value(key) = setting
		    
		    Exception err As RegistryAccessErrorException
		      // Just bail out
		      Return
		      
		  #Endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Second() As Integer
		  // Return the current second
		  Return Xojo.Core.Date.Now.Second
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub SendKeys(keys As String, unused As Boolean = False)
		  #If TargetWindows
		    // We want to initialize all of our virtual keys.  Some of them
		    // are going to be constants, others will be figured out while
		    // we parse, and still others will reside in a map.
		    Const VK_SHIFT = &h10
		    Const VK_CONTROL = &h11
		    Const VK_MENU = &h12
		    Const VK_ENTER = &h0D
		    
		    Static sMap As Dictionary
		    
		    If sMap = Nil Then
		      // Create our map
		      sMap = New Dictionary
		      
		      // Fill the map out (note that this is ByRef)
		      FillKeyMap(sMap)
		    End If
		    
		    // We have to write a very simple parser to parse
		    // the keys string that's passed in.
		    
		    // Split the entire string into a set of tokens.  Each token
		    // is a single character.
		    Dim chars() As String = Split(keys, "")
		    
		    // This holds the current virtual key (so we don't have to
		    // process the ASCII character multiple times).
		    Dim virtKey As Integer
		    
		    // This holds the current set of depressed modifier keys.
		    // As we press the keys, we should be adding to this list, 
		    // and when we need to release the keys, we know which 
		    // ones to generate a key up for.
		    Dim modifierList() As Integer
		    
		    // Sometimes we want to hold the modifiers for a while, for
		    // instance, if the user enters +(EC), then we want to hold
		    // the shift key for E and C.  But other times, we only want the
		    // modifier held for one key press, such as +EC, in which case
		    // there would be a Shift+E, and then a C.
		    Dim holdModifiers As Boolean
		    
		    // We may need to do some special processing for things inside
		    // of a {} tag, such as {TAB} or {LEFT 42}
		    Dim specialProcessing As Boolean
		    Dim specialProcessingStr As String
		    
		    // Loop over every token and do the appropriate
		    // action.
		    For Each token As String In chars
		      // If we're doing special processing, then we
		      // should be doing that instead of the normal
		      // processing.
		      If specialProcessing Then
		        // Check to see whether we have a { or
		        // not.  If we have one, and we've already
		        // processed at least one character, then we
		        // are done with our special processing.  The
		        // reason we check to see whether we have a
		        // character or not is because of the string {}}, 
		        // which should print a } character.
		        If token = "}" And Len(specialProcessingStr) > 0 Then
		          specialProcessing = False
		        Else
		          // Add the character to our processing string
		          specialProcessingStr = specialProcessingStr + token
		        End If
		      End If
		      
		      // If we're still doing special processing, then we want
		      // to continue doing it.  But since the state may have changed
		      // we need to check again
		      If specialProcessing Then Continue
		      
		      Select Case token
		      Case "("
		        // We're starting a token group.  The group
		        // should have the current modifiers applied 
		        // to it.
		        holdModifiers = True
		        
		      Case ")"
		        // If we're not holding modifies, then
		        // that means we've never gotten a ( and
		        // something is wrong
		        If Not holdModifiers Then Return
		        
		        // We're ending a token group.  This means 
		        // that we can clear the modifiers.
		        For Each modifier As Integer In modifierList
		          // Release the key
		          KeyUp(modifier)
		        Next modifier
		        
		        // Now clear our list
		        Redim modifierList(-1)
		        
		        // We're no longer holding the modifiers
		        holdModifiers = False
		        
		      Case "{"
		        // We have a special token to parse, such as
		        // {TAB} or {F1}.  We should search until we
		        // get to } and figure out what key to press
		        // from there.
		        //
		        // This could also be a repeat modifier that is
		        // in the form {key number}.  If we find a space
		        // while parsing, then we know we have this form.
		        
		        // So note that we need to parse until we hit the }.
		        // Once we have that, we can do the processing.
		        specialProcessing = True
		        
		      Case "}"
		        // Our special processing is done now, so we should
		        // check the data we have in the string.  It could be 
		        // in the form "key number", or it could just be "key".
		        
		        // First, get the key from our map.  If we don't have
		        // the key in the map, then we should try a regular
		        // ASCII key.
		        Dim firstField As String = NthField(specialProcessingStr, " ", 1)
		        Dim keyToken As Integer = sMap.Lookup(firstField, -1)
		        If keyToken = -1 Then
		          keyToken = ASCIIToScanKey(firstField)
		        End If
		        
		        // If we're still in a bad state, then we should
		        // bail out.
		        If keyToken = 0 Then Return
		        
		        // We want our repeats to always be at least one, but 
		        // possibly more, if that field exists.
		        Dim numRepeats As Integer = Max(Val(NthField(specialProcessingStr, " ", 2)), 1)
		        For count As Integer = 1 To numRepeats
		          // Press the key and then release it
		          KeyDown(keyToken)
		          KeyUp(keyToken)
		        Next count
		        
		        // Clear out our special processing data
		        specialProcessingStr = ""
		      Case "~"
		        // We want to send an enter key
		        KeyDown(VK_ENTER)
		        KeyUp(VK_ENTER)
		        
		      Case "+"
		        // The shift key modifier should be pressed
		        If modifierList.IndexOf(VK_SHIFT) = -1 Then
		          KeyDown(VK_SHIFT)
		          modifierList.Append(VK_SHIFT)
		        End If
		        
		      Case "^"
		        // The control key modifier should be pressed
		        If modifierList.IndexOf(VK_CONTROL) = -1 Then
		          KeyDown(VK_CONTROL)
		          modifierList.Append(VK_CONTROL)
		        End If
		        
		      Case "%"
		        // The Alt key modifier should be pressed
		        If modifierList.IndexOf(VK_MENU) = -1 Then
		          KeyDown(VK_MENU)
		          modifierList.Append(VK_MENU)
		        End If
		        
		      Else
		        // We have a regular key press, such as A or 4.
		        virtKey = ASCIIToScanKey(token)
		        
		        // Check to see if the scan key we got back
		        // has any of the modifier keys pressed
		        // or not.
		        Dim releaseShift As Boolean
		        If Bitwise.BitAnd(virtKey, &h100) = &h100 Then
		          KeyDown(VK_SHIFT)
		          releaseShift = True
		        End If
		        
		        Dim releaseControl As Boolean
		        If Bitwise.BitAnd(virtKey, &h200) = &h200 Then
		          KeyDown(VK_CONTROL)
		          releaseControl = True
		        End If
		        
		        Dim releaseAlt As Boolean
		        If Bitwise.BitAnd(virtKey, &h400) = &h400 Then
		          KeyDown(VK_MENU)
		          releaseAlt = True
		        End If
		        
		        // Press the key and then release it
		        KeyDown(virtKey)
		        KeyUp(virtKey)
		        
		        If releaseAlt Then KeyUp(VK_MENU)
		        If releaseControl Then KeyUp(VK_CONTROL)
		        If releaseShift Then KeyUp(VK_SHIFT)
		        
		        // If we aren't holding modifiers, then we
		        // should release them all here.
		        If Not holdModifiers Then
		          For Each modifier As Integer In modifierList
		            // Release the key
		            KeyUp(modifier)
		          Next modifier
		          
		          // Now clear our list
		          Redim modifierList(-1)
		        End If
		      End Select
		    Next token
		  #Endif
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Sgn(number As Double) As Integer
		  If number = 0 Then Return 0
		  If number < 0 Then Return -1
		  If number > 0 Then Return 1
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Shell(pathname As String, style As Integer = 1) As Integer
		  // We want to launch the application given by the path name, and
		  // we need to return the application's PID if the launch was successful
		  #If TargetWin32
		    Soft Declare Function CreateProcessW Lib "Kernel32" (appName As Integer, params As WString, _
		    procAttribs As Integer, threadAttribs As Integer, inheritHandles As Boolean, flags As Integer, _
		    env As Integer, curDir As Integer, startupInfo As Ptr, procInfo As Ptr) As Boolean
		    
		    Soft Declare Function CreateProcessA Lib "Kernel32" (appName As Integer, params As CString, _
		    procAttribs As Integer, threadAttribs As Integer, inheritHandles As Boolean, flags As Integer, _
		    env As Integer, curDir As Integer, startupInfo As Ptr, procInfo As Ptr) As Boolean
		    
		    Dim startupInfo, procInfo As MemoryBlock
		    
		    startupInfo = New MemoryBlock(17 * 4)
		    procInfo = New MemoryBlock(16)
		    
		    Dim unicodeSavvy As Boolean = System.IsFunctionAvailable("CreateProcessW", "Kernel32")
		    
		    startupInfo.Long(0) = startupInfo.Size
		    
		    // Create the application
		    Dim ret As Boolean
		    If unicodeSavvy Then
		      ret = CreateProcessW(0, pathname, 0, 0, False, 0, 0, 0, startupInfo, procInfo)
		    Else
		      ret = CreateProcessA(0, pathname, 0, 0, False, 0, 0, 0, startupInfo, procInfo)
		    End If
		    
		    // If we couldn't make it, then we're stuck
		    If Not ret Then Return 0
		    
		    // We want to return the process identifier for the application
		    Dim retVal As Integer = procInfo.Long(8)
		    
		    // We should wait for the input idle so that we can switch to the app
		    Declare Function WaitForInputIdle Lib "User32" (handle As Integer, wait As Integer) As Integer
		    Dim wait As Integer = WaitForInputIdle(procInfo.Long(0), 2500)
		    
		    // Clean the application up
		    Declare Sub CloseHandle Lib "Kernel32" (handle As Integer)
		    CloseHandle(procInfo.Long(0))
		    CloseHandle(procInfo.Long(4))
		    
		    Return retVal
		  #EndIf
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Space(num as Integer) As String
		  // Return a string with the proper number
		  // of spaces.
		  
		  Return FillString(" ", num)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Sqr(d as Double) As Double
		  // This is the squareroot function, which Xojo has a different name for
		  Return Sqrt(d)
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Sub Stop()
		  // This behaves just like Break, so do that.
		  Break
		End Sub
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function StrReverse(s As String) As String
		  // Return s with the characters in reverse order.
		  
		  If Len(s) < 2 Then Return s
		  
		  Dim m As MemoryBlock
		  Dim c As String
		  Dim pos, mpos, csize As Integer
		  
		  m = NewMemoryBlock(s.LenB)
		  
		  pos = 1
		  mpos = m.Size
		  While mpos > 0
		    c = Mid(s, pos, 1)
		    csize = c.LenB
		    mpos = mpos - csize
		    m.StringValue(mpos, csize) = c
		    pos = pos + 1
		  Wend
		  
		  Return DefineEncoding(m.StringValue(0, m.Size), s.Encoding)
		  
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Time() As Xojo.Core.Date
		  Return Xojo.Core.Date.Now
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Timer() As Single
		  // This returns the number of seconds that have elapsed since
		  // midnight.
		  
		  // Get the current time
		  Dim d As New Date
		  
		  // Calculate the number of seconds since
		  // midnight.  We do this the cheap way.
		  Dim midnight As New Date
		  midnight.Hour = 0
		  
		  // Now that we have midnight and now, we
		  // can figure out the number of seconds by
		  // subtracting the totalseconds of each
		  Return d.TotalSeconds - midnight.TotalSeconds
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function UCase(s As String) As String
		  Return s.Uppercase
		End Function
	#tag EndMethod

	#tag Method, Flags = &h1
		Protected Function Year() As Integer
		  // Return the current year
		  Return Xojo.Core.Date.Now.Year
		End Function
	#tag EndMethod


	#tag ComputedProperty, Flags = &h21
		#tag Getter
			Get
			  Static r As New Random
			  Return r
			End Get
		#tag EndGetter
		Private mRnd As Random
	#tag EndComputedProperty


	#tag Constant, Name = kSettingsLocation, Type = String, Dynamic = False, Default = \"HKEY_CURRENT_USER\\Software\\VB and VBA Program Settings\\", Scope = Private
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
