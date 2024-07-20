using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using static Global;

//Load the workbook
XLWorkbook workBook = null;

String programState = UserInput();
if (programState == "1")
{
    workBook = LoadWorkbook("inventory.xlsx");
    //Initialize workbook and set global variables for quick access
    workBook = InitWorkbook(workBook);

    //Set categories for items on items worksheet
    workBook = ProcessCategories(workBook);

    //Sort the items by category for organization
    //Important to note is clover disregards extra columns 
    //when importing inventory, so this is no issue for
    //use with clover if we just leave it sorted
    workBook = SortWorkbookByCategory(workBook);

    //Create new sheet in workbook specifically for taking inventory
    //Just like extra columns, extra sheets do not cause issues with 
    //clovers inventory system.
    workBook = CreateInventorySheet(workBook);
    Console.WriteLine("Inventory sheet from CloverPOS has been processed");
    Console.WriteLine("You may now use Excel to count inventory on the inventory sheet");
    Console.WriteLine("Once your count has been finished, come back here and use option 2");
    Console.WriteLine("This will prepare an inventory sheet for you to upload to CLoverPOS");
    Console.WriteLine();
    Console.WriteLine();
    workBook = CreateInventoryValueSheet(workBook);
    SaveWorkbook(workBook, "output.xlsx");
    Console.WriteLine("inventory sheet has been saved as 'output.xlsx'");
}
else if (programState == "2")
{
    Console.WriteLine("Please input inventory sheets name. If left blank, 'output.xlsx' will be used");
    String ip = Console.ReadLine();
    if (ip == null || ip == "")
        workBook = LoadWorkbook("output.xlsx");
    else workBook = LoadWorkbook(ip);


    //Process finished inventory 
    workBook = ProcessCompletedInventory(workBook);
    Console.WriteLine("Inventory count completed");
    SaveWorkbook(workBook, "outputinventory_done.xlsx");
    Console.WriteLine("Inventory sheet for CloverPOS saved as 'inventory_done.xlsx");
}
else Environment.Exit(0);

XLWorkbook LoadWorkbook(String fileName)
{
    XLWorkbook wb = new XLWorkbook(fileName);

    return wb;
}

String UserInput()
{
    String answer;
    Console.WriteLine("This is a CloverPOS inventory helper");
    Console.WriteLine("You can either prepare a worksheet for counting inventory, ");
    Console.WriteLine("or you can prepare a workbook to upload to CloverPOS after");
    Console.WriteLine("inventory count has been completed on the inventory sheet");
    Console.WriteLine("Please select one of the following options");
    Console.WriteLine(" ");
    Console.WriteLine(" ");
    Console.WriteLine(" ");
    Console.WriteLine("1) Process clovers inventory workbook, and prepare for inventory count");
    Console.WriteLine("2) Process completed inventory count workbook");
    answer = Console.ReadLine();

    if(answer != "1" && answer != "2")
    {
        Console.WriteLine(answer + " is not a valid option. Restart app when ready.");
        answer = "0";
    }
        

    return answer;
}

XLWorkbook InitWorkbook(XLWorkbook wb)
{
    //Fetch item worksheet
    var ws = wb.Worksheet(itemSheetRef);
    //Find next empty column and create category column
    catColRef = ws.LastColumnUsed().ColumnNumber() + 1;
    ws.Column(catColRef).FirstCell().Value = "Category";
    foreach(var col in ws.ColumnsUsed())
    {
        col.Width = 12;
    }

    return wb;
}

XLWorkbook ProcessCategories(XLWorkbook wb)
{
    var itemWS = wb.Worksheet(itemSheetRef);//set up quick ref to our items worksheet
    var catWS = wb.Worksheet(catSheetRef);//set up quick ref to our category sheet

    //We are looking for what category every item is in (might be none, might be multiple)
    //Iterate through every item row in items work sheet except the header
    var itemRows = itemWS.RangeUsed().RowsUsed().Skip(1);
    foreach (var row in itemRows)
    {
        row.Cell(catColRef).Value = "";
        //Iterate through every column of category sheet except the first
        var catColumns = catWS.RangeUsed().ColumnsUsed().Skip(1);
        foreach (var col in catColumns)
        {
            //Iterate through every cell of the column actually looking for the item name
            //This is done because the column is the category and each cell has an item value
            foreach (var cell in col.CellsUsed().Skip(1))
            {
                //If the items name is equal to the cell name
                if (row.Cell(itemColRef).Value.ToString() == cell.Value.ToString())
                {
                    //add cell category to item category
                    row.Cell(catColRef).Value = row.Cell(catColRef).Value.ToString() + col.Cell(1).Value + ", ";
                }
            }
        }
    }
    return wb;
}

//Sort Item worksheet by category name
XLWorkbook SortWorkbookByCategory(XLWorkbook wb)
{
    var ws = wb.Worksheet(itemSheetRef);
    //Make temp copy of first row so it is excluded from sort
    List<String> tempCells = new List<String>();
    foreach (var cell in ws.Row(1).CellsUsed())
    {
        tempCells.Add(cell.Value.ToString());
    }
    //Remove row1 from existance
    ws.Row(1).Delete();


    //Sorting by column is easy
    ws.RangeUsed().Sort(catColRef);

    //Insert temp row at top of sheet
    ws.Row(1).InsertRowsAbove(1);
    //Copy temprow to 1st row 
    int i = 1;
    foreach (String cell in tempCells)
    {
        ws.Row(1).Cell(i).Value = cell;
        i++;
    }

    return wb;
}

XLWorkbook CreateInventorySheet(XLWorkbook wb)
{
    var ws = wb.AddWorksheet("Inventory Count");
    var itemws = wb.Worksheet(itemSheetRef);

    //Create first row of inventory worksheet
    ws.Row(1).Cell(1).Value = "Product";
    ws.Row(1).Cell(2).Value = "UPC";
    ws.Row(1).Cell(3).Value = "Quantity";
    ws.Row(1).Cell(4).Value = "Count";
    //Set Columm widths
    ws.Column(1).Width = 48;
    ws.Column(2).Width = 18;
    ws.Column(3).Width = 10;
    ws.Column(4).Width = 10;

    //Create rows with data from item worksheet
    int i = 2; //i starts at 2 because row1 is the header
    foreach (var row in itemws.RangeUsed().RowsUsed().Skip(1))
    {
        ws.Row(i).Cell(1).Value = row.Cell(itemColRef).Value;
        ws.Row(i).Cell(2).Value = row.Cell(upcColRef).Value;
        ws.Row(i).Cell(3).Value = row.Cell(quantityColRef).Value;
        ws.Row(i).Cell(4).Value = " ";//Left blank. Must fill out for inventory
        i++;
    }

    foreach(var row in ws.RangeUsed().RowsUsed())
    {
        //Put a border around cells for easy inventory
        row.Cell(1).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        row.Cell(2).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        row.Cell(3).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
        row.Cell(4).Style.Border.OutsideBorder = XLBorderStyleValues.Medium;
    }
    ws.Unhide();
    ws.Unprotect();
    return wb;
}

XLWorkbook ProcessCompletedInventory(XLWorkbook wb)
{
    Console.WriteLine("Workbook Loaded");

    //-----------------
    //Time to process
    //-----------------

    var itemWS = wb.Worksheet(itemSheetRef);//set up quick ref to our items worksheet
    var countWS = wb.Worksheet(5);//set up quick ref to our inventory count sheet

    //Iterate through every item row in items work sheet except the header
    var itemRows = itemWS.RangeUsed().RowsUsed().Skip(1);
    foreach (var row in itemRows)
    {
        //Iterate through every row of the count worksheet
        var countRows = countWS.RangeUsed().RowsUsed().Skip(1);
        foreach (var count in countRows)
        {
            //Compare name field of both rows. If it's a match we have our row to update quantity
            if(count.Cell(1).Value.ToString() == row.Cell(itemColRef).Value.ToString())
            {
                if (count.Cell(4).Value.ToString() != "" && count.Cell(4).Value.ToString() != " " && count.Cell(4).Value.ToString() != null)
                {
                    row.Cell(quantityColRef).Value = count.Cell(4).Value;
                }
                if (count.Cell(2).Value.ToString().Length >= 4)
                    row.Cell(upcColRef).Value = count.Cell(2).Value;

                break;
            }
        }
    }
    return wb;
}

//Create an inventory value assement
//Does not count negative quantity inventory items
XLWorkbook CreateInventoryValueSheet(XLWorkbook wb)
{
    var ws = wb.AddWorksheet("Inventory Value");

    //Set values of two rows in column 1
    ws.Row(1).Cell(1).Value = "Inventory Cost";
    ws.Row(2).Cell(1).Value = "Inventory Retail Value";

    double invCost = 0.0d;
    double invVal = 0.0d;
    foreach (var row in wb.Worksheet(1).RowsUsed().Skip(1))
    {
        
        var quantity = row.Cell(quantityColRef).Value.ToString();
        if (quantity.Length > 0)//if quantity has a value
        {
            if (int.Parse(quantity) >= 0)//if quantity is not negativ
            {
                var cost = row.Cell(costColRef).Value.ToString();
                if (cost.Length > 0)
                    invCost += Math.Round(double.Parse(cost) * int.Parse(quantity), 2);//Inventory cost += cost * quantity
                var val = row.Cell(priceColRef).Value.ToString();
                if (val.Length > 0)
                    invVal += Math.Round(double.Parse(val) * int.Parse(quantity), 2);
            }
        }
    }
    
    ws.Row(1).Cell(2).Value = invCost;
    ws.Row(2).Cell(2).Value = invVal;
    return wb;
}


void SaveWorkbook(XLWorkbook wb, String fileName)
{
    wb.SaveAs(fileName);

}

//Hard values
public static class Global
{
    //Worksheet for items
    public static int itemSheetRef = 1;
    //worksheet for categories
    public static int catSheetRef = 3;
    //Column for items in item worksheet
    public static int itemColRef = 2;
    //Column for categories in item worksheet
    public static int catColRef;
    //Column for UPCs in item worksheet
    public static int upcColRef = 9;
    //Column for Quantity on hand of item in item sheet
    public static int quantityColRef = 12;
    public static int costColRef = 8;
    public static int priceColRef = 4;
}