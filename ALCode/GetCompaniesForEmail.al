codeunit 60301 "NLOutlookExtension"
{
    var
        RegHelper: Codeunit "NL Registration Helper";
        SalesSetup: Record "Sales & Receivables Setup";

    [ServiceEnabled]
    procedure GetCompaniesForEmail(EmailAddress: Text[250]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        Customer: Record Customer;
        Companies: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
    begin
        // Search in Customer table
        Customer.SetRange("E-Mail", EmailAddress);
        if Customer.FindSet() then
            repeat
                Companies.Add(Format(Customer.Name) + ' (' + Customer."No." + ')');
            until Customer.Next() = 0;

        // Define pagination range
        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        if (PageNumber * PageSize) < Companies.Count() then
            EndIndex := PageNumber * PageSize
        else
            EndIndex := Companies.Count();

        // Generate response for the requested page
        for Index := StartIndex to EndIndex do
            Result += Companies.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    [ServiceEnabled]
    procedure GetProjectsForCompany(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer; IncludeClosedProjects: Boolean) returnValue: Text
    var
        Project: Record "Job";
        Companies: Record Customer;
        Projects: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
    begin
        // Find the company by ID
        if not Companies.Get(CompanyId) then
            Error('Company with ID "%1" not found.', CompanyId);

        // Find projects for the company
        Project.SetRange("Bill-to Customer No.", Companies."No.");
        if not IncludeClosedProjects then
            Project.SetRange(Status, Project.Status::Open); // Filter for open projects only

        if Project.FindSet() then
            repeat
                Projects.Add(Format(Project.Description) + ' (' + Project."No." + ')');
            until Project.Next() = 0;

        // Define pagination range
        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        if StartIndex + PageSize - 1 < Projects.Count() then
            EndIndex := StartIndex + PageSize - 1
        else
            EndIndex := Projects.Count();

        // Generate response for the requested page
        for Index := StartIndex to EndIndex do
            Result += Projects.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    [ServiceEnabled]
    procedure GetSalesQuotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        SalesHeader: Record "Sales Header";
        Quotes: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
    begin
        // Filter sales quotes for the specific company
        SalesHeader.SetRange("Sell-to Customer No.", CompanyId);
        SalesHeader.SetRange("Document Type", SalesHeader."Document Type"::Quote);

        if SalesHeader.FindSet() then
            repeat
                Quotes.Add(Format(SalesHeader."No.") + ': ' + Format(SalesHeader."Document Date"));
            until SalesHeader.Next() = 0;

        // Define pagination range
        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        if StartIndex + PageSize - 1 < Quotes.Count() then
            EndIndex := StartIndex + PageSize - 1
        else
            EndIndex := Quotes.Count();

        // Generate response
        for Index := StartIndex to EndIndex do
            Result += Quotes.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    [ServiceEnabled]
    procedure GetLastInvoices(CompanyId: Code[20]; Count: Integer) returnValue: Text
    var
        Invoice: Record "Sales Invoice Header";
        Invoices: List of [Text];
        Result: Text[1024];
        Index: Integer;
    begin
        // Filter invoices for the specific company
        Invoice.SetRange("Sell-to Customer No.", CompanyId);

        // Sort by posting date descending to get the latest invoices
        Invoice.SetCurrentKey("Posting Date");
        Invoice.Ascending(false);

        if Invoice.FindSet() then
            repeat
                Invoices.Add(Format(Invoice."No.") + ': ' + Format(Invoice."Posting Date"));
            until (Invoice.Next() = 0) or (Invoices.Count() >= Count);

        // Generate response
        for Index := 1 to Invoices.Count() do
            Result += Invoices.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    [ServiceEnabled]
    procedure GetCreditNotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        CreditMemo: Record "Sales Cr. Memo Entity Buffer";
        CreditNotes: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
    begin
        // Filter credit notes for the specific company
        CreditMemo.SetRange("Sell-to Customer No.", CompanyId);

        if CreditMemo.FindSet() then
            repeat
                CreditNotes.Add(Format(CreditMemo."No.") + ': ' + Format(CreditMemo."Posting Date"));
            until CreditMemo.Next() = 0;

        // Define pagination range
        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        if StartIndex + PageSize - 1 < CreditNotes.Count() then
            EndIndex := StartIndex + PageSize - 1
        else
            EndIndex := CreditNotes.Count();

        // Generate response
        for Index := StartIndex to EndIndex do
            Result += CreditNotes.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    [ServiceEnabled]
    procedure GetCompanyDetails(CompanyId: Code[20]) returnValue: Text
    var
        Customer: Record Customer;
        Contact: Record Contact;
        OtherContacts: Record Contact;
        JsonObject: JsonObject;
        JsonContactsArray: JsonArray;
        JsonContact: JsonObject;
        JsonString: Text;
    begin
        // Fetch company details
        if not Customer.Get(CompanyId) then
            Error('Company with ID "%1" not found.', CompanyId);

        // Add main company details to JSON object
        JsonObject.Add('CompanyName', Customer.Name);
        JsonObject.Add('Balance', Customer.Balance);
        JsonObject.Add('OverdueAmount', Customer."Balance Due");

        // Fetch main contact details
        if Contact.Get(Customer."Primary Contact No.") then begin
            Clear(JsonContact); // Reset the JsonObject for the primary contact
            JsonContact.Add('Name', Contact.Name);
            JsonContact.Add('Phone', Contact."Phone No.");
            JsonContact.Add('Email', Contact."E-Mail");
            JsonContact.Add('JobTitle', Contact."Job Title");
            JsonObject.Add('PrimaryContact', JsonContact);
        end;

        // Fetch other contacts associated with the company
        OtherContacts.SetRange("Company No.", Customer."No.");
        OtherContacts.SetRange(Type, OtherContacts.Type::Person); // Ensure only person contacts are fetched
        if OtherContacts.FindSet() then
            repeat
                Clear(JsonContact); // Reset the JsonObject for each new contact
                JsonContact.Add('Name', OtherContacts.Name);
                JsonContact.Add('Phone', OtherContacts."Phone No.");
                JsonContact.Add('Email', OtherContacts."E-Mail");
                JsonContact.Add('JobTitle', OtherContacts."Job Title");
                JsonContactsArray.Add(JsonContact);
            until OtherContacts.Next() = 0;

        // Add other contacts to the main JSON object
        JsonObject.Add('OtherContacts', JsonContactsArray);

        // Convert JSON object to string
        JsonObject.WriteTo(JsonString);

        // Return JSON string
        exit(JsonString);
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
    begin
        // Get customer number series
        if SalesSetup.Get() then begin
            CustNoSeriesCode := SalesSetup."Customer Nos.";
            if CustNoSeriesCode = '' then
                Error('No. Series for Customer Nos. is not defined in Sales & Receivables Setup.');
            CustNo := NoSeries.GetNextNo(CustNoSeriesCode, Today(), true);
        end else
            Error('Sales & Receivables Setup not found.');

        // Initialize and populate Customer record
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

        if not Customer.Insert() then begin
            result := 'Error';
            exit;
        end;

        // Apply customer template if needed
        TemplateCode := RegHelper.SelectCustomerTemplate(InvoiceLanguage, CountryCode);
        if TemplateCode <> '' then
            if CustomerTemplate.Get(TemplateCode) then begin
                CustomerTemplMgt.ApplyCustomerTemplate(Customer, CustomerTemplate, true);
                Customer.Modify();

                // Adjust Currency Code and Language Code if necessary
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

        // Get contact number series
        if MarketingSetup.Get() then begin
            ContactNoSeriesCode := MarketingSetup."Contact Nos.";
            if ContactNoSeriesCode = '' then
                Error('No. Series for Contact Nos. is not defined in Marketing Setup.');
            ContactNoForCompany := NoSeries.GetNextNo(ContactNoSeriesCode, Today(), true);
        end else
            Error('Marketing Setup not found.');

        // Create the primary contact
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

        if not Contact.Insert() then begin
            result := 'Error';
            exit;
        end;

        // Create the primary contact person
        PrimaryContactNo := RegHelper.NewContact(ContactNoForCompany, PrimaryContactFirstAndLastName, PrimaryContactDirectPhoneNumber, PrimaryContactMobilePhoneNumber, PrimaryContactEmail, PrimaryContactTitle);

        Customer.Validate("Primary Contact No.", PrimaryContactNo);
        Customer.Modify();

        // Return the customer ID (CustNo) instead of the contact number
        result := CustNo;
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
    var
        Customer: Record Customer;
        Contact: Record Contact;
        NeedsModify: Boolean;
    begin
        // Find the customer by ID
        if not Customer.Get(CustomerId) then
            Error('Customer with ID "%1" not found.', CustomerId);

        // Update customer fields
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

        // Apply modifications if necessary
        if NeedsModify then
            Customer.Modify();

        // Update the primary contact if available
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

        // Return success message
        returnValue := 'Customer with ID ' + CustomerId + ' updated successfully.';
    end;

    // ---------------------------MOCK DATA UNTIL EXTENSION ADDED TO SANDBOX-------------------------------
    [ServiceEnabled]
    procedure GetTaskAnalysis(TaskId: Code[20]) returnValue: Text
    var
        JsonObject: JsonObject;
        AnalysisObject: JsonObject;
        ChargeableObject: JsonObject;
        FreeOfChargeObject: JsonObject;
        JsonString: Text;
    begin
        // Mock Analysis Section
        AnalysisObject.Add('TotalWorked', 46.25);
        AnalysisObject.Add('TotalAdjustments', -2.25);
        AnalysisObject.Add('TotalFreeOfCharge', -2.00);
        AnalysisObject.Add('TotalChargeable', 42.00);
        AnalysisObject.Add('Total100PercentDiscount', 0.50);
        AnalysisObject.Add('TotalShownOnInvoice', 42.50);
        AnalysisObject.Add('ExpectedBilling', 0.50);

        // Mock Chargeable Section
        ChargeableObject.Add('Invoiced', 41.50);
        ChargeableObject.Add('Registered', 0.50);
        ChargeableObject.Add('TotalChargeable', 42.00);

        // Mock Free of Charge Section
        FreeOfChargeObject.Add('Internal', 1.50);
        FreeOfChargeObject.Add('100PercentDiscountInvoiced', 0.50);
        FreeOfChargeObject.Add('TotalFreeOfCharge', 2.00);

        // Combine all sections into the main JSON object
        JsonObject.Add('Analysis', AnalysisObject);
        JsonObject.Add('Chargeable', ChargeableObject);
        JsonObject.Add('FreeOfCharge', FreeOfChargeObject);

        // Convert JSON object to string
        JsonObject.WriteTo(JsonString);

        // Return the JSON response
        exit(JsonString);
    end;



}
