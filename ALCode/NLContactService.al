codeunit 60703 "NLContactService"
{

    SingleInstance = true;
    Access = Public;

    var
        RegHelper: Codeunit "NL Registration Helper";

    procedure AddOtherContacts(OtherContacts: Text; CustomerId: Code[20]): Text
    var
        Position: Integer;
        ContactFields: array[5] of Text;
        ContactLine: Text;
        RemainingContacts: Text;
        Contact: Record Contact;
        ContactNo: Code[20];
        NoSeries: Codeunit "No. Series";
        MarketingSetup: Record "Marketing Setup";
        ContactNoSeriesCode: Text[20];
        ContactBusinessRel: Record "Contact Business Relation";
        RegHelper: Codeunit "NL Registration Helper";
    begin
        OtherContacts := DelChr(OtherContacts, '<>', '"');
        RemainingContacts := OtherContacts;

        if MarketingSetup.Get() then begin
            ContactNoSeriesCode := MarketingSetup."Contact Nos.";
            if ContactNoSeriesCode = '' then
                Error('No. Series for Contact Nos. is not defined in Marketing Setup.');
        end else
            Error('Marketing Setup not found.');

        while RemainingContacts <> '' do begin
            Position := StrPos(RemainingContacts, '],');
            if Position = 0 then
                Position := StrLen(RemainingContacts);

            ContactLine := CopyStr(RemainingContacts, 2, Position - 2);
            RemainingContacts := DelStr(RemainingContacts, 1, Position + 1);

            if ContactLine <> '' then begin
                ContactFields[1] := SelectStr(1, ContactLine);
                ContactFields[2] := SelectStr(2, ContactLine);
                ContactFields[3] := SelectStr(3, ContactLine);
                ContactFields[4] := SelectStr(4, ContactLine);
                ContactFields[5] := DelChr(SelectStr(5, ContactLine), '<>', ']');

                ContactFields[3] := RegHelper.FilterOutLetters(ContactFields[3]);
                ContactFields[4] := RegHelper.FilterOutLetters(ContactFields[4]);

                if ContactFields[1] <> '' then begin
                    ContactNo := NoSeries.GetNextNo(ContactNoSeriesCode, Today(), true);

                    Contact.Init();
                    Contact."No." := ContactNo;
                    Contact.Type := Contact.Type::Person;
                    Contact."Company No." := CustomerId;
                    Contact.Validate(Name, ContactFields[1].Trim());
                    Contact.Validate("Job Title", ContactFields[2].Trim());
                    Contact.Validate("Phone No.", ContactFields[3].Trim());
                    Contact.Validate("Mobile Phone No.", ContactFields[4].Trim());
                    Contact.Validate("E-Mail", CopyStr(ContactFields[5].Trim(), 1, 80));

                    if not Contact.Insert() then
                        Error('Error creating contact for "%1".', ContactFields[1]);

                    if not ContactBusinessRel.Get(Contact."No.", ContactBusinessRel."Link to Table"::Customer, CustomerId) then begin
                        ContactBusinessRel.Init();
                        ContactBusinessRel.Validate("Contact No.", Contact."No.");
                        ContactBusinessRel.Validate("Link to Table", ContactBusinessRel."Link to Table"::Customer);
                        ContactBusinessRel.Validate("No.", CustomerId);
                        ContactBusinessRel.Insert();
                    end;
                end;
            end;
        end;

        exit('Other contacts added successfully.');
    end;

    procedure UpdateOtherContacts(CustomerId: Code[20]; UpdatedContacts: Text): Text
    var
        Contact: Record Contact;
        ContactFields: array[6] of Text;
        ContactLine: Text;
        NeedsModify: Boolean;
        ContactEntries: List of [Text];
        Entry: Text;
        RegHelper: Codeunit "NL Registration Helper";
    begin
        UpdatedContacts := CopyStr(UpdatedContacts, 2, StrLen(UpdatedContacts) - 2);
        ContactEntries := UpdatedContacts.Split('],[');

        foreach Entry in ContactEntries do begin
            ContactLine := DelChr(Entry, '<>', '[]');

            ContactFields[1] := DelChr(SelectStr(1, ContactLine), '<>', ' ');
            ContactFields[2] := DelChr(SelectStr(2, ContactLine), '<>', ' ');
            ContactFields[3] := DelChr(SelectStr(3, ContactLine), '<>', ' ');
            ContactFields[4] := DelChr(SelectStr(4, ContactLine), '<>', ' ');
            ContactFields[5] := DelChr(SelectStr(5, ContactLine), '<>', ' ');
            ContactFields[6] := DelChr(SelectStr(6, ContactLine), '<>', ' ');

            if not Contact.Get(ContactFields[1]) then
                Error('Contact "%1" not found.', ContactFields[1]);

            if Contact."Company No." <> CustomerId then
                Error('Contact "%1" does not belong to customer "%2".', ContactFields[1], CustomerId);

            NeedsModify := false;

            if ContactFields[2] <> '' then begin
                Contact.Validate(Name, ContactFields[2]);
                NeedsModify := true;
            end;

            if ContactFields[3] <> '' then begin
                Contact.Validate("Job Title", ContactFields[3]);
                NeedsModify := true;
            end;

            if ContactFields[4] <> '' then begin
                Contact.Validate("Phone No.", RegHelper.FilterOutLetters(ContactFields[4]));
                NeedsModify := true;
            end;

            if ContactFields[5] <> '' then begin
                Contact.Validate("Mobile Phone No.", RegHelper.FilterOutLetters(ContactFields[5]));
                NeedsModify := true;
            end;

            if ContactFields[6] <> '' then begin
                RegHelper.ValidateEmail(ContactFields[6]);
                Contact.Validate("E-Mail", ContactFields[6]);
                NeedsModify := true;
            end;

            if NeedsModify then
                Contact.Modify();
        end;

        exit('Contacts updated successfully.');
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
    ): Code[20]
    var
        Contact: Record Contact;
        Customer: Record Customer;
        CustomerTemplate: Record "Customer Templ.";
        MarketingSetup: Record "Marketing Setup";
        CustomerTemplMgt: Codeunit "Customer Templ. Mgt.";
        NoSeries: Codeunit "No. Series";
        NeedsModify: Boolean;
        ContactNoForCompany: Code[20];
        CustNo: Code[20];
        PrimaryContactNo: Code[20];
        TemplateCode: Code[20];
        ContactNoSeriesCode: Text[20];
        CustNoSeriesCode: Text[20];
        SalesSetup: Record "Sales & Receivables Setup";
    begin
        if SalesSetup.Get() then begin
            CustNoSeriesCode := SalesSetup."Customer Nos.";
            if CustNoSeriesCode = '' then
                Error('No. Series for Customer Nos. is not defined in Sales & Receivables Setup.');
            CustNo := NoSeries.GetNextNo(CustNoSeriesCode, Today(), true);
        end else
            Error('Sales & Receivables Setup not found.');

        Customer.Init();
        Customer."No." := CustNo;
        Customer.Validate(Name, CompanyName);
        Customer.Validate(Address, Address);
        Customer.Validate("Address 2", Address2);
        Customer.Validate("Post Code", PostalCode);
        Customer.Validate(City, City);
        Customer.Validate("Country/Region Code", CountryCode);
        Customer.Validate("VAT Registration No.", Cvr);
        Customer.Validate("Phone No.", PhoneNumber);
        RegHelper.ValidateEmail(InvoiceEmail);
        Customer."E-Mail" := InvoiceEmail;

        if not Customer.Insert() then
            Error('Error inserting customer record.');

        TemplateCode := RegHelper.SelectCustomerTemplate(InvoiceLanguage, CountryCode);
        if TemplateCode <> '' then
            if CustomerTemplate.Get(TemplateCode) then begin
                CustomerTemplMgt.ApplyCustomerTemplate(Customer, CustomerTemplate, true);
                Customer.Modify();
                NeedsModify := false;
                if CustomerTemplate."Currency Code" <> InvoiceCurrency then begin
                    Customer.Validate("Currency Code", InvoiceCurrency);
                    NeedsModify := true;
                end;
                if CustomerTemplate."Language Code" <> InvoiceLanguage then begin
                    Customer.Validate("Language Code", InvoiceLanguage);
                    NeedsModify := true;
                end;
                if NeedsModify then
                    Customer.Modify();
            end;

        if MarketingSetup.Get() then begin
            ContactNoSeriesCode := MarketingSetup."Contact Nos.";
            if ContactNoSeriesCode = '' then
                Error('No. Series for Contact Nos. is not defined in Marketing Setup.');
            ContactNoForCompany := NoSeries.GetNextNo(ContactNoSeriesCode, Today(), true);
        end else
            Error('Marketing Setup not found.');

        Contact.Init();
        Contact."No." := ContactNoForCompany;
        Contact.Type := Contact.Type::Company;
        Contact."Company No." := Customer."No.";
        Contact.Validate(Name, CompanyName);
        Contact.Validate(Address, Address);
        Contact.Validate("Address 2", Address2);
        Contact.Validate("Post Code", PostalCode);
        Contact.Validate(City, City);
        RegHelper.ValidateEmail(InvoiceEmail);
        Contact.Validate("Phone No.", PhoneNumber);
        Contact."E-Mail" := InvoiceEmail;

        if not Contact.Insert() then
            Error('Error inserting primary company contact.');

        PrimaryContactNo := RegHelper.NewContact(ContactNoForCompany, PrimaryContactFirstAndLastName, PrimaryContactDirectPhoneNumber, PrimaryContactMobilePhoneNumber, PrimaryContactEmail, PrimaryContactTitle);

        Customer.Validate("Primary Contact No.", PrimaryContactNo);
        Customer.Modify();

        exit(CustNo);
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
    ): Text
    var
        Customer: Record Customer;
        Contact: Record Contact;
        NeedsModify: Boolean;
    begin
        if not Customer.Get(CustomerId) then
            Error('Customer with ID "%1" not found.', CustomerId);

        NeedsModify := false;
        if CompanyName <> '' then begin
            Customer.Validate(Name, CompanyName);
            NeedsModify := true;
        end;
        if Address <> '' then begin
            Customer.Validate(Address, Address);
            NeedsModify := true;
        end;
        if Address2 <> '' then begin
            Customer.Validate("Address 2", Address2);
            NeedsModify := true;
        end;
        if PostalCode <> '' then begin
            Customer.Validate("Post Code", PostalCode);
            NeedsModify := true;
        end;
        if City <> '' then begin
            Customer.Validate(City, City);
            NeedsModify := true;
        end;
        if CountryCode <> '' then begin
            Customer.Validate("Country/Region Code", CountryCode);
            NeedsModify := true;
        end;
        if Cvr <> '' then begin
            Customer.Validate("VAT Registration No.", Cvr);
            NeedsModify := true;
        end;
        if PhoneNumber <> '' then begin
            Customer.Validate("Phone No.", PhoneNumber);
            NeedsModify := true;
        end;
        if InvoiceEmail <> '' then begin
            RegHelper.ValidateEmail(InvoiceEmail);
            Customer."E-Mail" := InvoiceEmail;
            NeedsModify := true;
        end;
        if InvoiceLanguage <> '' then begin
            Customer.Validate("Language Code", InvoiceLanguage);
            NeedsModify := true;
        end;
        if InvoiceCurrency <> '' then begin
            Customer.Validate("Currency Code", InvoiceCurrency);
            NeedsModify := true;
        end;
        if NeedsModify then
            Customer.Modify();

        if Customer."Primary Contact No." <> '' then begin
            if Contact.Get(Customer."Primary Contact No.") then begin
                if PrimaryContactFirstAndLastName <> '' then
                    Contact.Validate(Name, PrimaryContactFirstAndLastName);
                if PrimaryContactDirectPhoneNumber <> '' then
                    Contact.Validate("Phone No.", PrimaryContactDirectPhoneNumber);
                if PrimaryContactMobilePhoneNumber <> '' then
                    Contact.Validate("Mobile Phone No.", PrimaryContactMobilePhoneNumber);
                if PrimaryContactEmail <> '' then begin
                    RegHelper.ValidateEmail(PrimaryContactEmail);
                    Contact.Validate("E-Mail", PrimaryContactEmail);
                end;
                if PrimaryContactTitle <> '' then
                    Contact.Validate("Job Title", PrimaryContactTitle);
                Contact.Modify();
            end else
                Error('Primary contact for customer "%1" not found.', CustomerId);
        end;

        exit('Customer with ID ' + CustomerId + ' updated successfully.');
    end;
}