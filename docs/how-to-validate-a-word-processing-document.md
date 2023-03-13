# Validate a word processing document (Open XML SDK)

This topic shows how to use the classes in the Open XML SDK 2.5 for
Office to programmatically validate a word processing document.

The following assembly directives are required to compile the code in
this topic.

```csharp
    using System;
    using DocumentFormat.OpenXml.Packaging;
    using DocumentFormat.OpenXml.Validation;
    using DocumentFormat.OpenXml.Wordprocessing;
```



--------------------------------------------------------------------------------
## How the Sample Code Works
This code example consists of two methods. The first method, **ValidateWordDocument**, is used to validate a
regular Word file. It doesn't throw any exceptions and closes the file
after running the validation check. The second method, **ValidateCorruptedWordDocument**, starts by
inserting some text into the body, which causes a schema error. It then
validates the Word file, in which case the method throws an exception on
trying to open the corrupted file. The validation is done by using the
[Validate](https://msdn.microsoft.com/library/office/documentformat.openxml.validation.openxmlvalidator.validate.aspx) method. The code displays
information about any errors that are found, in addition to the count of
errors.


--------------------------------------------------------------------------------
## Sample Code
In your main method, you can call the two methods, **ValidateWordDocument** and **ValidateCorruptedWordDocument** by using the
following example that validates a file named "Word18.docx.".

```csharp
    string filepath = @"C:\Users\Public\Documents\Word18.docx";
    ValidateWordDocument(filepath);
    Console.WriteLine("The file is valid so far.");
    Console.WriteLine("Inserting some text into the body that would cause Schema error");
    Console.ReadKey();

    ValidateCorruptedWordDocument(filepath);
    Console.WriteLine("All done! Press a key.");
    Console.ReadKey();
```



> [!Important] 
> Notice that you cannot run the code twice after corrupting the file in the first run. You have to start with a new Word file.

Following is the complete sample code in both C\# and Visual Basic.

```csharp
    public static void ValidateWordDocument(string filepath)
    {
        using (WordprocessingDocument wordprocessingDocument =
        WordprocessingDocument.Open(filepath, true))
        {                  
            try
            {           
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in
                    validator.Validate(wordprocessingDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("ErrorType: " + error.ErrorType);
                    Console.WriteLine("Node: " + error.Node);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }

                Console.WriteLine("count={0}", count);
                }
                
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);              
            }

            wordprocessingDocument.Close();
        }
    }

    public static void ValidateCorruptedWordDocument(string filepath)
    {
        // Insert some text into the body, this would cause Schema Error
        using (WordprocessingDocument wordprocessingDocument =
        WordprocessingDocument.Open(filepath, true))
        {
            // Insert some text into the body, this would cause Schema Error
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            Run run = new Run(new Text("some text"));
            body.Append(run);

            try
            {
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                foreach (ValidationErrorInfo error in
                    validator.Validate(wordprocessingDocument))
                {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("ErrorType: " + error.ErrorType);
                    Console.WriteLine("Node: " + error.Node);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }

                Console.WriteLine("count={0}", count);
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
```



--------------------------------------------------------------------------------
## See also


- [Open XML SDK 2.5 class library reference](/office/open-xml/open-xml-sdk)
