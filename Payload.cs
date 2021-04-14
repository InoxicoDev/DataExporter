using System;

namespace DataExporter
{
    public class Payload
    {
        public string InoxicoReference { get; set; }
        public Data Data { get; set; }
        public object EndpointResponse { get; set; }
        public Integrationauditinfo IntegrationAuditInfo { get; set; }
        public Inputmodel InputModel { get; set; }
    }

    public class Data
    {
        public Companydetails CompanyDetails { get; set; }
        public Instruction Instruction { get; set; }
        public Validationoutcome ValidationOutcome { get; set; }
        public string _Type { get; set; }
    }

    public class Companydetails
    {
        public string NoxId { get; set; }
        public string EntityName { get; set; }
        public string TradingAsName { get; set; }
        public string EntityRegistrationNumber { get; set; }
        public string VatNumber { get; set; }
        public string EntityType { get; set; }
        public Taxdetails TaxDetails { get; set; }
        public Physicaladdress PhysicalAddress { get; set; }
        public Postaladdress PostalAddress { get; set; }
        public Bbbeedetails BBBEEDetails { get; set; }
        public Productsandservice[] ProductsAndServices { get; set; }
        public Contactperson ContactPerson { get; set; }
        public string Municipality { get; set; }
        public string _Type { get; set; }
    }

    public class Taxdetails
    {
        public string TaxNumber { get; set; }
        public string DocumentURL { get; set; }
        public string _Type { get; set; }
    }

    public class Physicaladdress
    {
        public string StreetAddress1 { get; set; }
        public string StreetAddress2 { get; set; }
        public string PostalCode { get; set; }
        public string City { get; set; }
        public string Province { get; set; }
        public string Country { get; set; }
        public string _Type { get; set; }

        public override string ToString()
        {
            return
                $"{StreetAddress1}; {StreetAddress2}; Postal Code: {PostalCode}; City: {City}; State/Province: {Province}; Country: {Country}";
        }
    }

    public class Postaladdress
    {
        public string StreetAddress1 { get; set; }
        public string StreetAddress2 { get; set; }
        public string PostalCode { get; set; }
        public string City { get; set; }
        public string Province { get; set; }
        public string Country { get; set; }
        public string _Type { get; set; }
    }

    public class Bbbeedetails
    {
        public string DocumentURL { get; set; }
        public DateTime CertificateExpiryDate { get; set; }
        public string Issuer { get; set; }
        public int Level { get; set; }
        public string BlackOwnershipPercentage { get; set; }
        public string BlackWomenOwnershipPercentage { get; set; }
        public string DocumentType { get; set; }
        public string TurnoverDuringAccreditation { get; set; }
        public string EnterpriseSize { get; set; }
        public string DesignatedGroup { get; set; }
        public bool EmpoweringSupplier { get; set; }
        public string _Type { get; set; }
    }

    public class Contactperson
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string ContactNumber { get; set; }
        public string CellPhoneNumber { get; set; }
        public string _Type { get; set; }
    }

    public class Productsandservice
    {
        public string CommodityClass { get; set; }
        public string Description { get; set; }
        public string _Type { get; set; }
    }

    public class Instruction
    {
        public string UpdateType { get; set; }
        public string UpdateStatus { get; set; }
        public DateTime ActionDate { get; set; }
        public string RFQNumbers { get; set; }
        public string TransactionReference { get; set; }
        public string _Type { get; set; }
    }

    public class Validationoutcome
    {
        public string OverallValidation { get; set; }
        public string RegistrationNumber { get; set; }
        public string CompanyStatus { get; set; }
        public string CompanyName { get; set; }
        public string VatNumber { get; set; }
        public string TaxNumber { get; set; }
        public string SubjectEmail { get; set; }
        public string BBBEEStructure { get; set; }
        public string BBBEEIssuer { get; set; }
        public string AccountVerification { get; set; }
        public string _Type { get; set; }
    }

    public class Integrationauditinfo
    {
        public string ClientWorkflowInstanceId { get; set; }
        public string InitiatingStepInstanceId { get; set; }
        public string SubjectNoxId { get; set; }
        public string ClientId { get; set; }
        public string _Type { get; set; }
    }

    public class Inputmodel
    {
        public string InoxicoReference { get; set; }
        public Submissiondata SubmissionData { get; set; }
        public Integrationauditingdata IntegrationAuditingData { get; set; }
        public string _Type { get; set; }
    }

    public class Submissiondata
    {
        public Companydetails1 CompanyDetails { get; set; }
        public Instruction1 Instruction { get; set; }
        public Validationoutcome1 ValidationOutcome { get; set; }
        public string _Type { get; set; }
    }

    public class Companydetails1
    {
        public string NoxId { get; set; }
        public string EntityName { get; set; }
        public string TradingAsName { get; set; }
        public string EntityRegistrationNumber { get; set; }
        public string VatNumber { get; set; }
        public string EntityType { get; set; }
        public Taxdetails1 TaxDetails { get; set; }
        public Physicaladdress1 PhysicalAddress { get; set; }
        public Postaladdress1 PostalAddress { get; set; }
        public Bbbeedetails1 BBBEEDetails { get; set; }
        public Productsandservice1[] ProductsAndServices { get; set; }
        public Contactperson1 ContactPerson { get; set; }
        public string Municipality { get; set; }
        public string _Type { get; set; }
    }

    public class Taxdetails1
    {
        public string TaxNumber { get; set; }
        public string DocumentURL { get; set; }
        public string _Type { get; set; }
    }

    public class Physicaladdress1
    {
        public string StreetAddress1 { get; set; }
        public string StreetAddress2 { get; set; }
        public string PostalCode { get; set; }
        public string City { get; set; }
        public string Province { get; set; }
        public string Country { get; set; }
        public string _Type { get; set; }
    }

    public class Postaladdress1
    {
        public string StreetAddress1 { get; set; }
        public string StreetAddress2 { get; set; }
        public string PostalCode { get; set; }
        public string City { get; set; }
        public string Province { get; set; }
        public string Country { get; set; }
        public string _Type { get; set; }
    }

    public class Bbbeedetails1
    {
        public string DocumentURL { get; set; }
        public DateTime CertificateExpiryDate { get; set; }
        public string Issuer { get; set; }
        public int Level { get; set; }
        public string BlackOwnershipPercentage { get; set; }
        public string BlackWomenOwnershipPercentage { get; set; }
        public string DocumentType { get; set; }
        public string TurnoverDuringAccreditation { get; set; }
        public string EnterpriseSize { get; set; }
        public string DesignatedGroup { get; set; }
        public bool EmpoweringSupplier { get; set; }
        public string _Type { get; set; }
    }

    public class Contactperson1
    {
        public string Name { get; set; }
        public string Email { get; set; }
        public string ContactNumber { get; set; }
        public string CellPhoneNumber { get; set; }
        public string _Type { get; set; }
    }

    public class Productsandservice1
    {
        public string CommodityClass { get; set; }
        public string Description { get; set; }
        public string _Type { get; set; }
    }

    public class Instruction1
    {
        public string UpdateType { get; set; }
        public string UpdateStatus { get; set; }
        public DateTime ActionDate { get; set; }
        public string RFQNumbers { get; set; }
        public string TransactionReference { get; set; }
        public string _Type { get; set; }
    }

    public class Validationoutcome1
    {
        public string OverallValidation { get; set; }
        public string RegistrationNumber { get; set; }
        public string CompanyStatus { get; set; }
        public string CompanyName { get; set; }
        public string VatNumber { get; set; }
        public string TaxNumber { get; set; }
        public string SubjectEmail { get; set; }
        public string BBBEEStructure { get; set; }
        public string BBBEEIssuer { get; set; }
        public string AccountVerification { get; set; }
        public string _Type { get; set; }
    }

    public class Integrationauditingdata
    {
        public string ClientWorkflowInstanceId { get; set; }
        public string InitiatingStepInstanceId { get; set; }
        public string SubjectNoxId { get; set; }
        public string ClientId { get; set; }
        public string _Type { get; set; }
    }
}
