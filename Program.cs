using System;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Excel = Microsoft.Office.Interop.Excel;


namespace DataExporter
{
    class Program
    {
        static void Main(string[] args)
        {
            Workbook wb = new Workbook(@"C:\Users\zandreb\Desktop\ExxaroExtracts\DecompressedExxaroExtracts.xlsx");
            Worksheet worksheet = wb.Worksheets[0];
            Cells cells = worksheet.Cells;
            // Get row and column count
            int rowCount = cells.MaxDataRow;



            // Used to export all in a single file
            ExportAllToSingleWorkbook(wb, worksheet, cells, rowCount);

            // Used when exporting each application in detail to it's own excell file
            //for (int row = 5954; row <= rowCount; row++)
            //{
            //    try
            //    {
            //        var subjectCompanyName = Convert.ToString(cells[row, 3].Value);
            //        var subjectCompanyNoxId = Convert.ToString(cells[row, 7].Value);
            //        var requestReference = Convert.ToString(cells[row, 15].Value);
            //        var jsonOutput = Convert.ToString(cells[row, 21].Value);

            //        var converted = JsonSerializer.Deserialize<Payload>(jsonOutput);

            //        ExportWorkbook($"{subjectCompanyName}_{subjectCompanyNoxId}_{requestReference}", converted);

            //        Console.WriteLine($"Book {row} exported");
            //    }
            //    catch (Exception e)
            //    {
            //        Console.WriteLine($"Error in book {row}");
            //    }
            //}
        }

        public static void ExportAllToSingleWorkbook(Workbook wb, Worksheet worksheet, Cells cells, int rowCount)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                throw new Exception("not installed");
            }

            Excel.Workbook xlWorkBook;
            object misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            var sheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);

            sheet.Name = "IntegrationData";
            sheet.Cells[1, 1] = "CompanyName";
            sheet.Cells[1, 2] = "RegistrationNumber";
            sheet.Cells[1, 3] = "TaxVatNumber";
            sheet.Cells[1, 4] = "EnterpriseSize";
            sheet.Cells[1, 5] = "BBBEEInformation_BlackOwnership";
            sheet.Cells[1, 6] = "BBBEEInformation_BlackWomanOwnership";
            sheet.Cells[1, 7] = "BBBEELevel";
            sheet.Cells[1, 8] = "PhysicalAddress";
            sheet.Cells[1, 9] = "Municipality";
            sheet.Cells[1, 10] = "ContactName";
            sheet.Cells[1, 11] = "ContactEmailAddress";
            sheet.Cells[1, 12] = "ContactTelephoneNumber";
            sheet.Cells[1, 13] = "ContactCellphoneNumber";
            sheet.Cells[1, 14] = "RFQNumbers";
            sheet.Cells[1, 15] = "LightValidationOutcome";
            sheet.Cells[1, 15] = "JSONIntegrationData";

            for (int row = 1; row <= rowCount; row++)
            {
                try
                {
                    var subjectCompanyName = Convert.ToString(cells[row, 3].Value);
                    var subjectCompanyNoxId = Convert.ToString(cells[row, 7].Value);
                    var requestReference = Convert.ToString(cells[row, 15].Value);
                    var jsonOutput = Convert.ToString(cells[row, 21].Value);

                    var converted = JsonSerializer.Deserialize<Payload>(jsonOutput);

                    sheet.Cells[row + 1, 1] = converted.Data.CompanyDetails.EntityName;
                    sheet.Cells[row + 1, 2] = converted.Data.CompanyDetails.EntityRegistrationNumber;
                    sheet.Cells[row + 1, 3] = converted.Data.CompanyDetails.TaxDetails.TaxNumber;
                    sheet.Cells[row + 1, 4] = converted.Data.CompanyDetails.BBBEEDetails.EnterpriseSize;
                    sheet.Cells[row + 1, 5] = converted.Data.CompanyDetails.BBBEEDetails.BlackOwnershipPercentage;
                    sheet.Cells[row + 1, 6] = converted.Data.CompanyDetails.BBBEEDetails.BlackWomenOwnershipPercentage; ;
                    sheet.Cells[row + 1, 7] = converted.Data.CompanyDetails.BBBEEDetails.Level; ;
                    sheet.Cells[row + 1, 8] = converted.Data.CompanyDetails.PhysicalAddress.ToString();
                    sheet.Cells[row + 1, 9] = converted.Data.CompanyDetails.Municipality;
                    sheet.Cells[row + 1, 10] = converted.Data.CompanyDetails.ContactPerson.Name;
                    sheet.Cells[row + 1, 11] = converted.Data.CompanyDetails.ContactPerson.Email;
                    sheet.Cells[row + 1, 12] = converted.Data.CompanyDetails.ContactPerson.ContactNumber;
                    sheet.Cells[row + 1, 13] = converted.Data.CompanyDetails.ContactPerson.CellPhoneNumber;
                    sheet.Cells[row + 1, 14] = converted.Data.Instruction.RFQNumbers;
                    sheet.Cells[row + 1, 15] = converted.Data.ValidationOutcome.OverallValidation;

                    Console.WriteLine($"Book {row} exported");
                }
                catch (Exception e)
                {
                    sheet.Cells[row + 1, 15] = cells[row, 21].Value;
                }
            }

            var file = @$"C:\Users\zandreb\Desktop\ExxaroExtracts\SingleFileExtract\IntegrationData_AllSubjectCompanies.xlsx";
            xlWorkBook.SaveAs(file);
            xlWorkBook.Close();

            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        public static void ExportWorkbook(string fileName, Payload payload)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                throw new Exception("not installed");
            }

            Excel.Workbook xlWorkBook;

            object misValue = System.Reflection.Missing.Value;
            var xlWorkbooks = xlApp.Workbooks;
            xlWorkBook = xlWorkbooks.Add(misValue);

            var xlSheets = xlWorkBook.Sheets as Excel.Sheets;
            var companyDetailSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
            var taxDetailSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[2], Type.Missing, Type.Missing, Type.Missing);
            var physicalAddressSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[3], Type.Missing, Type.Missing, Type.Missing);
            var postalAddressSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[4], Type.Missing, Type.Missing, Type.Missing);
            var beeDetailsSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[5], Type.Missing, Type.Missing, Type.Missing);
            var porductsAndServicesSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[6], Type.Missing, Type.Missing, Type.Missing);
            var contactPersonSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[7], Type.Missing, Type.Missing, Type.Missing);
            var instructionSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[8], Type.Missing, Type.Missing, Type.Missing);
            var validationOutcomeSheet = (Excel.Worksheet)xlSheets.Add(xlSheets[9], Type.Missing, Type.Missing, Type.Missing);


            AddCompanyDetails(companyDetailSheet, payload.Data.CompanyDetails);
            AddTaxDetails(taxDetailSheet, payload.Data.CompanyDetails.TaxDetails);
            AddPhysicalAddress(physicalAddressSheet, payload.Data.CompanyDetails.PhysicalAddress);
            AddPostalAddress(postalAddressSheet, payload.Data.CompanyDetails.PostalAddress);
            AddBEEDetails(beeDetailsSheet, payload.Data.CompanyDetails.BBBEEDetails);
            AddProductsAndServices(porductsAndServicesSheet, payload.Data.CompanyDetails.ProductsAndServices);
            AddContactPerson(contactPersonSheet, payload.Data.CompanyDetails.ContactPerson);
            AddInstruction(instructionSheet, payload.Data.Instruction);
            AddValidationOutcome(validationOutcomeSheet, payload.Data.ValidationOutcome);

            var file = @$"C:\Users\zandreb\Desktop\ExxaroExtracts\Extracted\IntegrationData_{Regex.Replace(fileName, @"\s+", "")}.xlsx";
            xlWorkBook.SaveAs(file);
            xlWorkBook.Close();

            xlApp.Quit();

            Marshal.FinalReleaseComObject(companyDetailSheet);
            Marshal.FinalReleaseComObject(taxDetailSheet);
            Marshal.FinalReleaseComObject(physicalAddressSheet);
            Marshal.FinalReleaseComObject(postalAddressSheet);
            Marshal.FinalReleaseComObject(beeDetailsSheet);
            Marshal.FinalReleaseComObject(porductsAndServicesSheet);
            Marshal.FinalReleaseComObject(contactPersonSheet);
            Marshal.FinalReleaseComObject(instructionSheet);
            Marshal.FinalReleaseComObject(validationOutcomeSheet);
            Marshal.FinalReleaseComObject(xlSheets);
            Marshal.FinalReleaseComObject(xlWorkBook);
            Marshal.FinalReleaseComObject(xlWorkbooks);
            Marshal.FinalReleaseComObject(xlApp);
        }

        private static void AddCompanyDetails(Excel.Worksheet sheet, Companydetails companydetails)
        {
            sheet.Name = "CompanyDetails";
            sheet.Cells[1, 1] = "NoxId";
            sheet.Cells[1, 2] = "EntityName";
            sheet.Cells[1, 3] = "RegistrationNumber";
            sheet.Cells[1, 4] = "VatNumber";
            sheet.Cells[1, 5] = "EntityType";
            sheet.Cells[1, 6] = "Municipality";

            sheet.Cells[2, 1] = companydetails.NoxId;
            sheet.Cells[2, 2] = companydetails.EntityName;
            sheet.Cells[2, 3] = companydetails.EntityRegistrationNumber;
            sheet.Cells[2, 4] = companydetails.VatNumber;
            sheet.Cells[2, 5] = companydetails.EntityType;
            sheet.Cells[2, 6] = companydetails.Municipality;
        }

        private static void AddTaxDetails(Excel.Worksheet sheet, Taxdetails taxDetails)
        {
            sheet.Name = "TaxDetails";
            sheet.Cells[1, 1] = "TaxNumber";
            sheet.Cells[1, 2] = "DocumentUrl";

            sheet.Cells[2, 1] = taxDetails.TaxNumber;
            sheet.Cells[2, 2] = taxDetails.DocumentURL;
        }

        private static void AddPhysicalAddress(Excel.Worksheet sheet, Physicaladdress address)
        {
            sheet.Name = "PhysicalAddres";
            sheet.Cells[1, 1] = "StreetAddress1";
            sheet.Cells[1, 2] = "StreetAddress2";
            sheet.Cells[1, 3] = "PostalCode";
            sheet.Cells[1, 4] = "City";
            sheet.Cells[1, 5] = "Province";
            sheet.Cells[1, 6] = "Country";

            sheet.Cells[2, 1] = address.StreetAddress1;
            sheet.Cells[2, 2] = address.StreetAddress2;
            sheet.Cells[2, 3] = address.PostalCode;
            sheet.Cells[2, 4] = address.City;
            sheet.Cells[2, 5] = address.Province;
            sheet.Cells[2, 6] = address.Country;
        }

        private static void AddPostalAddress(Excel.Worksheet sheet, Postaladdress address)
        {
            sheet.Name = "PostalAddress";
            sheet.Cells[1, 1] = "StreetAddress1";
            sheet.Cells[1, 2] = "StreetAddress2";
            sheet.Cells[1, 3] = "PostalCode";
            sheet.Cells[1, 4] = "City";
            sheet.Cells[1, 5] = "Province";
            sheet.Cells[1, 6] = "Country";

            sheet.Cells[2, 1] = address.StreetAddress1;
            sheet.Cells[2, 2] = address.StreetAddress2;
            sheet.Cells[2, 3] = address.PostalCode;
            sheet.Cells[2, 4] = address.City;
            sheet.Cells[2, 5] = address.Province;
            sheet.Cells[2, 6] = address.Country;
        }

        private static void AddBEEDetails(Excel.Worksheet sheet, Bbbeedetails bbbeedetails)
        {
            sheet.Name = "BBBEEDetails";
            sheet.Cells[1, 1] = "DocumentUrl";
            sheet.Cells[1, 2] = "CertifiedExpiryDate";
            sheet.Cells[1, 3] = "Issuer";
            sheet.Cells[1, 4] = "Level";
            sheet.Cells[1, 5] = "BlackOwnershipPercentage";
            sheet.Cells[1, 6] = "BlackWomenOwnershipPercentage";
            sheet.Cells[1, 7] = "TurnoverDuringAccreditation";
            sheet.Cells[1, 8] = "EnterpriseSize";
            sheet.Cells[1, 9] = "DesignatedGroup";
            sheet.Cells[1, 10] = "EmpoweringSupplier";

            sheet.Cells[2, 1] = bbbeedetails.DocumentURL;
            sheet.Cells[2, 2] = bbbeedetails.CertificateExpiryDate;
            sheet.Cells[2, 3] = bbbeedetails.Issuer;
            sheet.Cells[2, 4] = bbbeedetails.Level;
            sheet.Cells[2, 5] = bbbeedetails.BlackOwnershipPercentage;
            sheet.Cells[2, 6] = bbbeedetails.BlackWomenOwnershipPercentage;
            sheet.Cells[2, 7] = bbbeedetails.TurnoverDuringAccreditation;
            sheet.Cells[2, 8] = bbbeedetails.EnterpriseSize;
            sheet.Cells[2, 9] = bbbeedetails.DesignatedGroup;
            sheet.Cells[2, 10] = bbbeedetails.EmpoweringSupplier;
        }

        private static void AddProductsAndServices(Excel.Worksheet sheet, Productsandservice[] productsandservices)
        {
            sheet.Name = "ProductsAndServicesDetails";
            sheet.Cells[1, 1] = "CommodityClass";
            sheet.Cells[1, 2] = "Description";

            var counter = 2;
            foreach (var p in productsandservices)
            {
                sheet.Cells[counter, 1] = p.CommodityClass;
                sheet.Cells[counter, 2] = p.Description;

                counter++;
            }
        }

        private static void AddContactPerson(Excel.Worksheet sheet, Contactperson contactPerson)
        {
            sheet.Name = "ContactPerson";
            sheet.Cells[1, 1] = "Name";
            sheet.Cells[1, 2] = "Email";
            sheet.Cells[1, 3] = "ContactNumber";
            sheet.Cells[1, 4] = "CellPhoneNumber";

            sheet.Cells[2, 1] = contactPerson.Name;
            sheet.Cells[2, 2] = contactPerson.Email;
            sheet.Cells[2, 3] = contactPerson.ContactNumber;
            sheet.Cells[2, 4] = contactPerson.CellPhoneNumber;
        }

        private static void AddInstruction(Excel.Worksheet sheet, Instruction instruction)
        {
            sheet.Name = "Instruction";
            sheet.Cells[1, 1] = "UpdateType";
            sheet.Cells[1, 2] = "UpdateStatus";
            sheet.Cells[1, 3] = "ActionDate";
            sheet.Cells[1, 4] = "RFQNumbers";
            sheet.Cells[1, 5] = "TransactionReference";

            sheet.Cells[2, 1] = instruction.UpdateType;
            sheet.Cells[2, 2] = instruction.UpdateStatus;
            sheet.Cells[2, 3] = instruction.ActionDate;
            sheet.Cells[2, 4] = instruction.RFQNumbers;
            sheet.Cells[2, 5] = instruction.TransactionReference;
        }

        private static void AddValidationOutcome(Excel.Worksheet sheet, Validationoutcome validationOutcome)
        {
            sheet.Name = "ValidationOutcome";
            sheet.Cells[1, 1] = "OverallValidation";
            sheet.Cells[1, 2] = "RegistrationNumber";
            sheet.Cells[1, 3] = "CompanyStatus";
            sheet.Cells[1, 4] = "CompanyName";
            sheet.Cells[1, 5] = "VatNumber";
            sheet.Cells[1, 6] = "TaxNumber";
            sheet.Cells[1, 7] = "SubjectEmail";
            sheet.Cells[1, 8] = "BBBEEStructure";
            sheet.Cells[1, 9] = "BBBEEIssuer";
            sheet.Cells[1, 10] = "AccountVerification";

            sheet.Cells[2, 1] = validationOutcome.OverallValidation;
            sheet.Cells[2, 2] = validationOutcome.RegistrationNumber;
            sheet.Cells[2, 3] = validationOutcome.CompanyStatus;
            sheet.Cells[2, 4] = validationOutcome.CompanyName;
            sheet.Cells[2, 5] = validationOutcome.VatNumber;
            sheet.Cells[2, 6] = validationOutcome.TaxNumber;
            sheet.Cells[2, 7] = validationOutcome.SubjectEmail;
            sheet.Cells[2, 8] = validationOutcome.BBBEEStructure;
            sheet.Cells[2, 9] = validationOutcome.BBBEEIssuer;
            sheet.Cells[2, 10] = validationOutcome.AccountVerification;
        }
    }
}
