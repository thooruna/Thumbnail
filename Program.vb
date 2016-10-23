'
' Created by SharpDevelop.
' User: Wilfred
' Date: 10/22/2016
' Time: 12:05 PM
' 
Imports System
Imports System.Drawing
Imports Inventor
Imports Microsoft.VisualBasic

Module Program
	
	Sub Main()
		Dim arguments As String() = System.Environment.GetCommandLineArgs()
		
		'Dim strFullFileName As String = System.IO.Path.GetTempPath + "_thumb.png"
	    Dim strFullFileName As String = "C:\temp\_thumb.png"
	    If System.IO.File.Exists(strFullFileName) Then
	    	System.IO.File.Delete(strFullFileName)
	    End If
	    
		If arguments.Length <> 2 Then
			Console.WriteLine("Invalid number of parameters")
		Else
			Dim strArgument As String
			strArgument = arguments(1)
			
			If strArgument = "/?" Then
				Console.WriteLine("No help defined")
			ElseIf System.IO.File.Exists(strArgument) Then			
				
				Dim objApp As New ApprenticeServerComponent
	    		Dim objDoc As ApprenticeServerDocument = objApp.Open(strArgument)
	    		Dim objThumb As stdole.IPictureDisp = objDoc.PropertySets.Item("Inventor Summary Information").Item("Thumbnail").value
	    				
	    		Dim objImage As System.Drawing.Image = New Bitmap(Compatibility.VB6.IPictureDispToImage(objThumb), New System.Drawing.Size(267,267))
	    		objImage.Save(strFullFileName)
	    			    		
        		objDoc.Close
        		
        		Console.WriteLine("Thumbnail created for: " + strArgument)
        		Console.WriteLine("Thumbnail saved as: " + strFullFileName)
			Else
				Console.WriteLine("The system cannot find the file specified: " + strArgument)
			End If
		End If
		
		'Console.Write("Press any key to continue . . . ")
		'Console.ReadKey(True)
	End Sub
	
End Module
