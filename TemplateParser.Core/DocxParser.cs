using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;         // Needed for WordProcessingDocument.
using DocumentFormat.OpenXml.Wordprocessing; 
using System.Collections.Generic;
using System.Text.Json;
namespace TemplateParser.Core;

public sealed class DocxParser
{
    public ParserResult ParseDocxTemplate(string filePath, Guid templateId)
    {
        //Creates a list to write down the 
        List<Node> listNodes = new List<Node>();
        //Creates a stack to know whos the parent of the next section
        Stack<(Node node, int level)> parentStack = new Stack<(Node node, int level)>();

        // Opens the word document in read mode.
        using (WordprocessingDocument wordProcessingDocument = WordprocessingDocument.Open(filePath, false))
        {
        
            //Opens the main part of the document and returns null is anything is null
            Body? body = wordProcessingDocument?.MainDocumentPart?.Document?.Body;
            //if the body is empty it will throw an error "Document is empty"
            ArgumentNullException.ThrowIfNull(body, "Document is empty.");
        

            // Loops through every paragraph.
            foreach (Paragraph p in body.Descendants<Paragraph>())
        {
            //this extracts the style of the paragaph, "No Style" means its a normal paragraph
            string? style = p?.ParagraphProperties?.ParagraphStyleId?.Val ?? "No Style";

            /*
            // Extracst the actual text.
            string? text = p?.InnerText;

            //Prints the Style firs the the actual text
            Console.WriteLine($"Style: {style}");
            Console.WriteLine($"Text: {text}");

            // I added the following line to space the output out a little better:
            Console.WriteLine("--------------------------------");
            */

            //This is the heading level
            int heading = 0;

            //THis determines what heading level to give the node depending on the heading
            if(style == "Heading1")
                {
                    heading = 1;
                }
                else if(style == "Heading2")
                {
                    heading = 2;
                }
                else if(style == "Heading3")
                {
                    heading = 3;
                }
                else
                {
                    continue;
                }

            //Determines what parent the node has
            while(parentStack.Count >0 && parentStack.Peek().level > heading)
                {
                    parentStack.Pop(); //will remove the node from the stack until it reaches the same level or less
                }

                //this determines what kind of type is added to the node
                string styleType = "Section";

                //this sets styletype to the type for the node
                if(style == "Heading1")
                {
                    styleType = "section";
                }
                else if(style == "Heading2")
                {
                    styleType = "subsection";
                }
                else if(style == "Heading3")
                {
                    styleType = "subsubsection";
                }

                //This is where the actual code is written
                Node newNode = new Node
                {
                    Id = Guid.NewGuid(), //each node has a personal UUID
                    TemplateId = templateId, //Used templateId from the parameters ofParseDocxTemplate
                    ParentId = parentStack.Count > 0 ? parentStack.Peek().node.Id : null, //finds the Id of the parent Section using the stack
                    Type = styleType, //This is where its set either section, subsection, or subsubsection
                    Title = style, //This is wether it was heading 1, 2, or 3
                    OrderIndex = 0, //what index its on, basically how many sections were made before(its set to zero at the start then changes next)
                    MetadataJson = "{}"
                };
            
                //This is where the OrderIndex are counted
                int siblingCount = 0; // counts how many other of the same nodes are under the Section
                foreach (var n in listNodes)
                    {
                        if (n.ParentId == newNode.ParentId)
                        {
                            siblingCount++;
                        }
                    }

                    //sets OrderIndex from newNode to the correct index
                    newNode.OrderIndex = siblingCount;

                    //adds the newNode to the listNodes
                    listNodes.Add(newNode);
                    //adds the newNode and the heading as a tuple to the stack
                    parentStack.Push((newNode, heading));

                    string jsonString = JsonSerializer.Serialize(listNodes, new JsonSerializerOptions
                    {
                        WriteIndented = true
                    });

                    Console.WriteLine(jsonString);




        }
    }





        // TODO (Week 1-4): Implement core DOCX parsing here.
        // Recommended responsibilities for this method:
        // 1) [Week 1] Learn DOCX structure and print paragraphs from the document.
        // 2) [Week 2] Build section hierarchy using Word heading styles.
        // 3) [Week 3] Detect tables, lists, and images as structured content nodes.
        // 4) [Week 4] Add formatting heuristics for files missing heading styles.
        // 5) [Week 2-4] Create Node instances with:
        //    - Id: new Guid for each node
        //    - TemplateId: the templateId argument
        //    - ParentId: null for root nodes, set for child nodes
        //    - Type/Title/OrderIndex/MetadataJson based on parsed content
        // 6) [Week 4] Return ParserResult with Nodes in deterministic order.
        //
        // Helper guidance [Week 3-6]:
        // - YES, create helper classes if this method gets long or hard to read.
        // - Keep helpers inside TemplateParser.Core (for example, Parsing/ or Utilities/ folders).
        // - Keep this method as the high-level orchestration entry point.
        // - In Week 6, refactor large blocks from this method into focused helper classes.
        //
        // Do not place parsing logic in the CLI project; keep it in Core.
        throw new NotImplementedException("DOCX parsing is intentionally not implemented in this starter repository.");
    }
}
