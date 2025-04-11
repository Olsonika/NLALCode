codeunit 60701 "NLOutlookExtension"
{
    var
        CompanyService: Codeunit "NLCompanyService";
        ProjectService: Codeunit "NLProjectService";
        ContactService: Codeunit "NLContactService";
        SalesService: Codeunit "NLSalesService";

    [ServiceEnabled]
    procedure GetCompaniesForEmail(EmailAddress: Text[250]; PageSize: Integer; PageNumber: Integer) returnValue: Text

    begin
        returnValue := CompanyService.GetCompaniesForEmail(EmailAddress, PageSize, PageNumber);
    end;


    [ServiceEnabled]
    procedure GetCompanyDetails(CompanyId: Code[20]; IsCustomer: Boolean) returnValue: Text
    begin
        returnValue := CompanyService.GetCompanyDetails(CompanyId, IsCustomer);
    end;

    procedure GetCurrentCompanyName(): Text
    begin
        exit(CompanyService.GetCurrentCompanyName());
    end;

    [ServiceEnabled]
    procedure GetProjectsForCompany(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer; IncludeClosedProjects: Boolean) returnValue: Text
    begin
        returnValue := ProjectService.GetProjectsForCompany(CompanyId, PageSize, PageNumber, IncludeClosedProjects);
    end;

    [ServiceEnabled]
    procedure GetTasksForProject(ProjectId: Code[20]) returnValue: Text
    begin
        returnValue := ProjectService.GetTasksForProject(ProjectId);
    end;


    [ServiceEnabled]
    procedure GetSalesQuotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    begin
        returnValue := SalesService.GetSalesQuotes(CompanyId, PageSize, PageNumber);
    end;

    [ServiceEnabled]
    procedure GetLastInvoices(CompanyId: Code[20]; Count: Integer) returnValue: Text
    begin
        returnValue := SalesService.GetLastInvoices(CompanyId, Count);
    end;

    [ServiceEnabled]
    procedure GetCreditNotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    begin
        returnValue := SalesService.GetCreditNotes(CompanyId, PageSize, PageNumber);
    end;



    [ServiceEnabled]
    procedure CreateCustomer(
        CompanyName: Text;
        Address: Text;
        Address2: Text;
        PostalCode: Text;
        City: Text;
        Cvr: Text;
        PhoneNumber: Text;
        InvoiceEmail: Text[80];
        PrimaryContactFirstAndLastName: Text;
        PrimaryContactMobilePhoneNumber: Text;
        PrimaryContactDirectPhoneNumber: Text;
        PrimaryContactEmail: Text[80];
        PrimaryContactTitle: Text;
        CountryCode: Text;
        InvoiceLanguage: Text;
        InvoiceCurrency: Text
    ) result: Code[20]

    begin
        result := ContactService.CreateCustomer(
            CompanyName, Address, Address2, PostalCode, City, Cvr, PhoneNumber, InvoiceEmail,
            PrimaryContactFirstAndLastName, PrimaryContactMobilePhoneNumber, PrimaryContactDirectPhoneNumber,
            PrimaryContactEmail, PrimaryContactTitle, CountryCode, InvoiceLanguage, InvoiceCurrency
        );
    end;

    [ServiceEnabled]
    procedure UpdateCustomer(
        CustomerId: Code[20];
        CompanyName: Text;
        Address: Text;
        Address2: Text;
        PostalCode: Text;
        City: Text;
        Cvr: Text;
        PhoneNumber: Text;
        InvoiceEmail: Text[80];
        PrimaryContactFirstAndLastName: Text;
        PrimaryContactMobilePhoneNumber: Text;
        PrimaryContactDirectPhoneNumber: Text;
        PrimaryContactEmail: Text[80];
        PrimaryContactTitle: Text;
        CountryCode: Text;
        InvoiceLanguage: Text;
        InvoiceCurrency: Text
    ) returnValue: Text
    begin
        returnValue := ContactService.UpdateCustomer(
            CustomerId, CompanyName, Address, Address2, PostalCode, City, Cvr,
            PhoneNumber, InvoiceEmail, PrimaryContactFirstAndLastName, PrimaryContactMobilePhoneNumber,
            PrimaryContactDirectPhoneNumber, PrimaryContactEmail, PrimaryContactTitle,
            CountryCode, InvoiceLanguage, InvoiceCurrency
        );
    end;

    [ServiceEnabled]
    procedure UpdateOtherContacts(CustomerId: Code[20]; UpdatedContacts: Text) returnValue: Text
    begin
        returnValue := ContactService.UpdateOtherContacts(CustomerId, UpdatedContacts);
    end;

    [ServiceEnabled]
    procedure AddOtherContacts(OtherContacts: Text; CustomerId: Code[20]) returnValue: Text
    begin
        returnValue := ContactService.AddOtherContacts(OtherContacts, CustomerId);
    end;


    procedure GetSalesQuoteLink(QuoteNo: Code[20]): Text
    begin
        exit(SalesService.GetSalesQuoteLink(QuoteNo));
    end;

    procedure GetInvoiceLink(InvoiceNo: Code[20]): Text
    begin
        exit(SalesService.GetInvoiceLink(InvoiceNo));
    end;

    procedure GetCreditNoteLink(CreditNoteNo: Code[20]): Text
    begin
        exit(SalesService.GetCreditNoteLink(CreditNoteNo));
    end;

    procedure EncodeDataString(Input: Text): Text
    var
        EncodedText: Text;
    begin
        EncodedText := Input;
        EncodedText := StrSubstNo(EncodedText, '%1', '%20'); // Replace spaces with %20
        EncodedText := StrSubstNo(EncodedText, '%', '%25');  // Replace % with %25
        EncodedText := StrSubstNo(EncodedText, '/', '%2F');  // Replace / with %2F
        EncodedText := StrSubstNo(EncodedText, '&', '%26');  // Replace & with %26
        EncodedText := StrSubstNo(EncodedText, '=', '%3D');  // Replace = with %3D
        EncodedText := StrSubstNo(EncodedText, '#', '%23');  // Replace # with %23
        EncodedText := StrSubstNo(EncodedText, '?', '%3F');  // Replace ? with %3F
        EncodedText := StrSubstNo(EncodedText, '\', '%5C'); // Replace \ with %5C
        exit(EncodedText);
    end;

    [ServiceEnabled]
    procedure GetTaskAnalysis(TaskId: Code[20]) returnValue: Text
    begin
        returnValue := ProjectService.GetTaskAnalysis(TaskId);
    end;
}
