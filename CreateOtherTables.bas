Option Compare Database
'Option Explicit

Sub Step00_Main()

Call Step01_CreateTables

End Sub

Sub Step01_CreateTables()

Dim strSQL As String

strSQL = "CREATE TABLE Customer("
strSQL = strSQL & "  CustomerID INT NOT NULL,"
strSQL = strSQL & "  Customer INT NOT NULL,"
strSQL = strSQL & "  CustomerPhone INT NOT NULL,"
strSQL = strSQL & "  Contact INT NOT NULL,"
strSQL = strSQL & "  Email INT NOT NULL,"
strSQL = strSQL & "  CurrentDiscount INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (CustomerID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE ArtworkOrder_Project("
strSQL = strSQL & "  ArtworkOrderID INT NOT NULL,"
strSQL = strSQL & "  OrderDate INT NOT NULL,"
strSQL = strSQL & "  MaximumColors INT NOT NULL,"
strSQL = strSQL & "  ArtWorkDescription INT NOT NULL,"
strSQL = strSQL & "  ArtworkFees INT NOT NULL,"
strSQL = strSQL & "  FixedCharges INT NOT NULL,"
strSQL = strSQL & "  ShippingCosts INT NOT NULL,"
strSQL = strSQL & "  DateApproved INT NOT NULL,"
strSQL = strSQL & "  ScheduledPrintDate INT NOT NULL,"
strSQL = strSQL & "  TotalPrice INT NOT NULL,"
strSQL = strSQL & "  CustomerID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (ArtworkOrderID),"
strSQL = strSQL & "  FOREIGN KEY (CustomerID) REFERENCES Customer(CustomerID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE PrintOrder("
strSQL = strSQL & "  PrintOrderID INT NOT NULL,"
strSQL = strSQL & "  PrintDate INT NOT NULL,"
strSQL = strSQL & "  ApparelOrderDate INT NOT NULL,"
strSQL = strSQL & "  Art_FilmDate INT NOT NULL,"
strSQL = strSQL & "  DateDelivered INT NOT NULL,"
strSQL = strSQL & "  PrintOrderDate INT NOT NULL,"
strSQL = strSQL & "  DueDate INT NOT NULL,"
strSQL = strSQL & "  Art_SlideDate INT NOT NULL,"
strSQL = strSQL & "  SetUpCharge INT NOT NULL,"
strSQL = strSQL & "  Deposit INT NOT NULL,"
strSQL = strSQL & "  Discount INT NOT NULL,"
strSQL = strSQL & "  ArtworkOrderID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (PrintOrderID),"
strSQL = strSQL & "  FOREIGN KEY (ArtworkOrderID) REFERENCES ArtworkOrder_Project(ArtworkOrderID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE Apparel_Item("
strSQL = strSQL & "  Vendor INT NOT NULL,"
strSQL = strSQL & "  TotalBlankPrice INT NOT NULL,"
strSQL = strSQL & "  Apparel_ID INT NOT NULL,"
strSQL = strSQL & "  BasePricePerUnit INT NOT NULL,"
strSQL = strSQL & "  PrintOrderID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (Apparel_ID),"
strSQL = strSQL & "  FOREIGN KEY (PrintOrderID) REFERENCES PrintOrder(PrintOrderID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE Size("
strSQL = strSQL & "  SizeCode INT NOT NULL,"
strSQL = strSQL & "  StandardPrice INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (SizeCode)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE Color("
strSQL = strSQL & "  ColorID INT NOT NULL,"
strSQL = strSQL & "  ColorDescription INT NOT NULL,"
strSQL = strSQL & "  ColorCharge INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (ColorID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE ArtWorkLocation("
strSQL = strSQL & "  ArtWorkLocationID INT NOT NULL,"
strSQL = strSQL & "  ArtWorkOrderID INT NOT NULL,"
strSQL = strSQL & "  LocationID INT NOT NULL,"
strSQL = strSQL & "  Description INT NOT NULL,"
strSQL = strSQL & "  Cost INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (ArtWorkLocationID),"
strSQL = strSQL & "  FOREIGN KEY (ArtworkOrderID) REFERENCES ArtworkOrder_Project(ArtworkOrderID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE Task("
strSQL = strSQL & "  TaskID INT NOT NULL,"
strSQL = strSQL & "  TaskName INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (TaskID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE Employee("
strSQL = strSQL & "  EmployeeID INT NOT NULL,"
strSQL = strSQL & "  EmployeeName INT NOT NULL,"
strSQL = strSQL & "  WageRate INT NOT NULL,"
strSQL = strSQL & "  EmployeePhone INT NOT NULL,"
strSQL = strSQL & "  FullTimeY_N INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (EmployeeID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE Material("
strSQL = strSQL & "  Material_ItemID INT NOT NULL,"
strSQL = strSQL & "  PriceCharged INT NOT NULL,"
strSQL = strSQL & "  PerUnitCost INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (Material_ItemID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE has("
strSQL = strSQL & "  has_ID INT NOT NULL,"
strSQL = strSQL & "  AdditionalCharge INT NOT NULL,"
strSQL = strSQL & "  Number_Units INT NOT NULL,"
strSQL = strSQL & "  Apparel_ID INT NOT NULL,"
strSQL = strSQL & "  SizeCode INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (has_ID),"
strSQL = strSQL & "  FOREIGN KEY (Apparel_ID) REFERENCES Apparel_Item(Apparel_ID),"
strSQL = strSQL & "  FOREIGN KEY (SizeCode) REFERENCES Size(SizeCode)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE has_base_color("
strSQL = strSQL & "  has_base_color_ID INT NOT NULL,"
strSQL = strSQL & "  ColorID INT NOT NULL,"
strSQL = strSQL & "  Apparel_ID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (has_base_color_ID),"
strSQL = strSQL & "  FOREIGN KEY (ColorID) REFERENCES Color(ColorID),"
strSQL = strSQL & "  FOREIGN KEY (Apparel_ID) REFERENCES Apparel_Item(Apparel_ID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE uses("
strSQL = strSQL & "  uses_ID INT NOT NULL,"
strSQL = strSQL & "  ArtworkOrderID INT NOT NULL,"
strSQL = strSQL & "  ColorID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (uses_ID),"
strSQL = strSQL & "  FOREIGN KEY (ArtWorkOrderID) REFERENCES ArtworkOrder_Project(ArtworkOrderID),"
strSQL = strSQL & "  FOREIGN KEY (ColorID) REFERENCES Color(ColorID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE includes("
strSQL = strSQL & "  includes_ID INT NOT NULL,"
strSQL = strSQL & "  MaterialUnits INT NOT NULL,"
strSQL = strSQL & "  Cost INT NOT NULL,"
strSQL = strSQL & "  Revenue INT NOT NULL,"
strSQL = strSQL & "  ArtworkOrderID INT NOT NULL,"
strSQL = strSQL & "  Material_ItemID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (includes_ID),"
strSQL = strSQL & "  FOREIGN KEY (ArtworkOrderID) REFERENCES ArtworkOrder_Project(ArtworkOrderID),"
strSQL = strSQL & "  FOREIGN KEY (Material_ItemID) REFERENCES Material(Material_ItemID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE EmployeeWorkLog("
strSQL = strSQL & "  WorkLogID INT NOT NULL,"
strSQL = strSQL & "  StartDate INT NOT NULL,"
strSQL = strSQL & "  TimeSpent INT NOT NULL,"
strSQL = strSQL & "  StartTime INT NOT NULL,"
strSQL = strSQL & "  ArtworkOrderID INT NOT NULL,"
strSQL = strSQL & "  EmployeeID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (WorkLogID),"
strSQL = strSQL & "  FOREIGN KEY (ArtworkOrderID) REFERENCES ArtworkOrder_Project(ArtworkOrderID),"
strSQL = strSQL & "  FOREIGN KEY (EmployeeID) REFERENCES Employee(EmployeeID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

strSQL = "CREATE TABLE contains("
strSQL = strSQL & "  contains_ID INT NOT NULL,"
strSQL = strSQL & "  WorkLogID INT NOT NULL,"
strSQL = strSQL & "  TaskID INT NOT NULL,"
strSQL = strSQL & "  PRIMARY KEY (contains_ID),"
strSQL = strSQL & "  FOREIGN KEY (WorkLogID) REFERENCES EmployeeWorkLog(WorkLogID),"
strSQL = strSQL & "  FOREIGN KEY (TaskID) REFERENCES Task(TaskID)"
strSQL = strSQL & ");"

Debug.Print strSQL
CurrentDb.Execute strSQL

Debug.Print "***Step01_CreateTables***"

End Sub