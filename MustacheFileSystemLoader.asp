<%
Class Mustache_Filesystem_Loader
	'''
    ' This class allows the use of different file partials.  http://mustache.github.com/
    ' To use:  loader = new Mustache_Filesystem_Loader
    '   	   loader.load('admin/dashboard') loads "./views/admin/dashboard.mustache"
    '''


    '''
    ''' Private Variables
    '''

    Private extension
	Private file_object
	Private cur_directory

    '''
    ' Mustache filesystem Loader constructor.
    '
	'	Gets the current directory the file is in
    '
    '''

    Public Sub Class_Initialize()
		cur_directory = Server.MapPath("/")
		Set file_object = Server.CreateObject("Scripting.FileSystemObject")
		extension = ".mustache"
    End Sub

    '''
    'Load a Template by name.
    '
	'	loader = new Mustache_Loader_FilesystemLoader()
    '   loader.load('admin/dashboard') loads "./views/admin/dashboard.mustache"
    '
    ' @param string name
    '
    ' @return string Mustache Template source
	'''
    Public Function load(name)
		Dim contents
        contents = loadFile(name)
		call cleanup
		load = contents
	End Function

    '''
    ' Helper function for loading a Mustache file by name.
    '
    ' @error string returns sError
    '
    ' @param string name
    '
    ' @return string Mustache Template contents
    '''
    Private Function loadFile(name)

        Dim read_file, file_name, contents, sError
		
		file_name = cur_directory & "\" & name & extension

        if not file_object.FileExists(file_name) then
            sError = "This file does not exist - " & file_name
			loadFile = sError
		else		
			Set read_file = file_object.OpenTextFile(file_name,1,false)
			contents =  read_file.ReadAll
			read_file.Close()
			loadFile = contents
		end if
    End Function

	'''
    ' Cleanup function
    '''
	Private Function cleanup()
		Set file_object = Nothing
	End Function
End Class
%>