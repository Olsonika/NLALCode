codeunit 60303 "NL Customer Service"
{
    procedure GetCompaniesForEmail(EmailAddress: Text[250]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        Customer: Record Customer;
        Companies: List of [Text];
        Result: Text[1024];
        Index, StartIndex, EndIndex : Integer;
    begin
        Customer.SetRange("E-Mail", EmailAddress);
        if Customer.FindSet() then
            repeat
                Companies.Add(Format(Customer.Name) + ' (' + Customer."No." + ')');
            until Customer.Next() = 0;

        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        EndIndex := PageNumber * PageSize;
        if EndIndex > Companies.Count() then
            EndIndex := Companies.Count();

        for Index := StartIndex to EndIndex do
            Result += Companies.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetCompanyDetails(CompanyId: Code[20]) returnValue: Text
    var
        Customer: Record Customer;
        JsonObject: JsonObject;
        JsonString: Text;
    begin
        if not Customer.Get(CompanyId) then
            Error('Company with ID "%1" not found.', CompanyId);

        JsonObject.Add('CompanyName', Customer.Name);
        JsonObject.Add('Balance', Customer.Balance);
        JsonObject.WriteTo(JsonString);
        exit(JsonString);
    end;

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
    var
        Customer: Record Customer;
        SalesSetup: Record "Sales & Receivables Setup";
        NoSeries: Codeunit "No. Series";
        CustomerNo: Code[20];
    begin
        SalesSetup.Get();
        CustomerNo := NoSeries.GetNextNo(SalesSetup."Customer Nos.", Today(), true);
        Customer.Init();
        Customer.Validate(Name, CompanyName);
        Customer.Validate(Address, Address);
        Customer.Validate("E-Mail", InvoiceEmail);
        Customer."No." := CustomerNo;
        Customer.Insert();

        exit(CustomerNo);
    end;

    procedure AddOtherContacts(OtherContacts: Text; CustomerId: Code[20])
    var
        Contact: Record Contact;
        ContactLine: Text;
        ContactFields: array[5] of Text;
        ContactNo: Code[20];
        NoSeries: Codeunit "No. Series";
    begin
        while OtherContacts <> '' do begin
            ContactLine := SelectStr(1, OtherContacts);
            OtherContacts := DelStr(OtherContacts, 1, StrPos(OtherContacts, ']') + 1);
            ContactFields[1] := SelectStr(1, ContactLine);
            ContactFields[2] := SelectStr(2, ContactLine);
            ContactFields[3] := SelectStr(3, ContactLine);
            ContactFields[4] := SelectStr(4, ContactLine);
            ContactFields[5] := SelectStr(5, ContactLine);

            ContactNo := NoSeries.GetNextNo('CONTACTNOS', Today(), true);
            Contact.Init();
            Contact."No." := ContactNo;
            Contact.Type := Contact.Type::Person;
            Contact."Company No." := CustomerId;
            Contact.Validate(Name, ContactFields[1]);
            Contact.Validate("Job Title", ContactFields[2]);
            Contact.Validate("Phone No.", ContactFields[3]);
            Contact.Validate("Mobile Phone No.", ContactFields[4]);
            Contact.Validate("E-Mail", ContactFields[5]);
            Contact.Insert();
        end;
    end;

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
    var
        Customer: Record Customer;
    begin
        if not Customer.Get(CustomerId) then
            Error('Customer with ID "%1" not found.', CustomerId);

        Customer.Validate(Name, CompanyName);
        Customer.Validate(Address, Address);
        Customer.Validate("E-Mail", InvoiceEmail);
        Customer.Modify();

        exit('Customer updated successfully.');
    end;

    procedure UpdateOtherContacts(CustomerId: Code[20]; UpdatedContacts: Text) returnValue: Text
    var
        Contact: Record Contact;
        ContactLine: Text;
        ContactFields: array[5] of Text;
    begin
        while UpdatedContacts <> '' do begin
            ContactLine := SelectStr(1, UpdatedContacts);
            UpdatedContacts := DelStr(UpdatedContacts, 1, StrPos(UpdatedContacts, ']') + 1);
            ContactFields[1] := SelectStr(1, ContactLine);
            ContactFields[2] := SelectStr(2, ContactLine);

            if Contact.Get(ContactFields[1]) then begin
                Contact.Validate(Name, ContactFields[2]);
                Contact.Modify();
            end;
        end;

        exit('Contacts updated successfully.');
    end;
}
