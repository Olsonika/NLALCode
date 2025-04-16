codeunit 60706 "NLCompanyService"
{
    SingleInstance = true;
    Access = Public;

    procedure GetCompaniesForEmail(EmailAddress: Text[250]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        Customer: Record Customer;
        Vendor: Record Vendor;
        Contact: Record Contact;
        Companies: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
        ContactName: Text[250];
        ContactEmail: Text[250];
    begin
        // Search in Customer table
        Customer.SetRange("E-Mail", EmailAddress);
        if Customer.FindSet() then
            repeat
                // Lookup primary contact for Customer
                if Contact.Get(Customer."Primary Contact No.") then begin
                    ContactName := Contact.Name;
                    ContactEmail := Contact."E-Mail";
                end else begin
                    ContactName := 'N/A';
                    ContactEmail := 'N/A';
                end;

                // Add Customer details
                Companies.Add(
                    Format(Customer.Name) + ' (' + Customer."No." + '), ' +
                    'Type: Customer, ' +
                    'Contact: ' + ContactName + ', ' +
                    'Email: ' + ContactEmail
                );
            until Customer.Next() = 0;

        // Search in Vendor table
        if Vendor.FindFirst() then
            repeat
                ContactName := 'N/A';
                ContactEmail := Vendor."E-Mail"; // Default to Vendor's Email field

                // Check if there is a linked Contact for the Vendor
                if Contact.Get(Vendor."Primary Contact No.") then begin
                    ContactName := Contact.Name;
                    ContactEmail := Contact."E-Mail"; // Use the Contact's Email
                end;

                // Only add the vendor if the email matches
                if ContactEmail = EmailAddress then
                    Companies.Add(
                        Format(Vendor.Name) + ' (' + Vendor."No." + '), ' +
                        'Type: Vendor, ' +
                        'Contact: ' + ContactName + ', ' +
                        'Email: ' + ContactEmail
                    );

            until Vendor.Next() = 0;

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
    procedure GetProjectsForCompany(CompanyId: Code[20]; IncludeClosedProjects: Boolean) returnValue: Text
    var
        Project: Record "Job";
        Companies: Record Customer;
        Projects: List of [Text];
        Result: Text[1024];
        BaseUrl: Text;
        ProjectLink: Text;
        ProjectStatus: Text;
        Index: Integer;
    begin
        // Base URL for project links
        BaseUrl := 'https://nl-server.navilogic.dk/bc24-intern';

        // Find the company by ID
        if not Companies.Get(CompanyId) then
            Error('Company with ID "%1" not found.', CompanyId);

        // Find projects for the company
        Project.SetRange("Bill-to Customer No.", Companies."No.");

        // Include all projects (both Open and Closed)
        if not IncludeClosedProjects then
            Project.SetRange(Status, Project.Status::Open);

        if Project.FindSet() then
            repeat
                // Determine Project Status
                if Project.Status = Project.Status::Open then
                    ProjectStatus := 'Open'
                else
                    ProjectStatus := 'Closed';

                // Generate a project link for page 88
                ProjectLink := StrSubstNo('%1?page=88&filter=''No.''%20IS%20''%2''', BaseUrl, Project."No.");

                // Add project with link **and status**
                Projects.Add(Format(Project.Description) + ' (' + Project."No." + ') [' + ProjectLink + '] - ' + ProjectStatus);
            until Project.Next() = 0;

        Index := 1;

        // Return all projects in a semicolon-separated format
        for Index := 1 to Projects.Count() do
            Result += Projects.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetCurrentCompanyName(): Text
    var
        CompanyInfo: Record "Company Information";
    begin
        if CompanyInfo.Get() then
            exit(CompanyInfo.Name)
        else
            Error('Company information is not available.');
    end;

    procedure GetCompanyDetails(CompanyId: Code[20]; IsCustomer: Boolean) returnValue: Text
    var
        Customer: Record Customer;
        Vendor: Record Vendor;
        Contact: Record Contact;
        OtherContacts: Record Contact;
        Salesperson: Record "Salesperson/Purchaser";
        JsonObject: JsonObject;
        JsonContactsArray: JsonArray;
        JsonContact: JsonObject;
        JsonString: Text;
        BaseUrl: Text;
        CompanyLink: Text;
        CreditLimit: Decimal;
        BlockedStatus: Text;
    begin
        // Base URL for the environment
        BaseUrl := 'https://nl-server.navilogic.dk/bc24-intern';

        if IsCustomer then begin
            // Fetch company details from the Customer table
            if not Customer.Get(CompanyId) then
                Error('Customer with ID "%1" not found.', CompanyId);

            // Calculate FlowFields
            Customer.CalcFields("Sales (LCY)", Balance, "Balance Due");

            // Generate company link
            CompanyLink := StrSubstNo('%1?page=21&filter=''No.''%20IS%20''%2''', BaseUrl, Customer."No.");

            // Prepare JSON data
            JsonObject.Add('CompanyType', 'Customer');
            JsonObject.Add('CompanyName', Customer.Name);
            JsonObject.Add('CompanyNo', Customer."No.");
            JsonObject.Add('CompanyLink', CompanyLink);
            JsonObject.Add('SalesLCY', Format(Customer."Sales (LCY)"));
            JsonObject.Add('BalanceLCY', Format(Customer.Balance));
            JsonObject.Add('BalanceDueLCY', Format(Customer."Balance Due"));

            // Add Credit Limit
            CreditLimit := Customer."Credit Limit (LCY)";
            if CreditLimit = 0 then
                JsonObject.Add('CreditLimitLCY', 'N/A')
            else
                JsonObject.Add('CreditLimitLCY', Format(CreditLimit));

            // Add general details
            JsonObject.Add('CVR', Customer."VAT Registration No.");
            JsonObject.Add('CountryCode', Customer."Country/Region Code");
            JsonObject.Add('PhoneNumber', Customer."Phone No.");
            JsonObject.Add('InvoiceEmail', Customer."E-Mail");

            // Add invoice language and currency
            JsonObject.Add('InvoiceLanguage', Customer."Language Code");
            JsonObject.Add('InvoiceCurrency', Customer."Currency Code");

            // Add blocked status
            BlockedStatus := Format(Customer.Blocked);
            if BlockedStatus = '' then
                BlockedStatus := 'No';
            JsonObject.Add('Blocked', BlockedStatus);

            // Add website
            JsonObject.Add('Website', Customer."Home Page");

            // Fetch salesperson details
            if Salesperson.Get(Customer."Salesperson Code") then
                JsonObject.Add('Salesperson', Salesperson.Name)
            else
                JsonObject.Add('Salesperson', 'N/A');

            // Add location details
            JsonObject.Add('Address', Customer.Address);
            JsonObject.Add('Address2', Customer."Address 2");
            JsonObject.Add('City', Customer.City);
            JsonObject.Add('PostalCode', Customer."Post Code");
            JsonObject.Add('Country', Customer."Country/Region Code");

            // Fetch main contact details
            if Contact.Get(Customer."Primary Contact No.") then begin
                Clear(JsonContact);
                JsonContact.Add('ContactNo', Contact."No.");
                JsonContact.Add('Name', Contact.Name);
                JsonContact.Add('Phone', Contact."Phone No.");
                JsonContact.Add('MobilePhone', Contact."Mobile Phone No.");
                JsonContact.Add('DirectPhone', Contact."Phone No.");
                JsonContact.Add('Email', Contact."E-Mail");
                JsonContact.Add('JobTitle', Contact."Job Title");
                JsonObject.Add('PrimaryContact', JsonContact);
            end;

            // Fetch other contacts
            OtherContacts.SetRange("Company No.", Customer."No.");
            OtherContacts.SetRange(Type, OtherContacts.Type::Person);
            if OtherContacts.FindSet() then
                repeat
                    Clear(JsonContact);
                    JsonContact.Add('ContactNo', OtherContacts."No.");
                    JsonContact.Add('Name', OtherContacts.Name);
                    JsonContact.Add('Phone', OtherContacts."Phone No.");
                    JsonContact.Add('MobilePhone', OtherContacts."Mobile Phone No.");
                    JsonContact.Add('DirectPhone', OtherContacts."Phone No.");
                    JsonContact.Add('Email', OtherContacts."E-Mail");
                    JsonContact.Add('JobTitle', OtherContacts."Job Title");
                    JsonContactsArray.Add(JsonContact);
                until OtherContacts.Next() = 0;

            // Add other contacts to the main JSON object
            JsonObject.Add('OtherContacts', JsonContactsArray);
        end else begin
            // Fetch company details from the Vendor table
            if not Vendor.Get(CompanyId) then
                Error('Vendor with ID "%1" not found.', CompanyId);

            // Generate company link
            CompanyLink := StrSubstNo('%1?page=26&filter=''No.''%20IS%20''%2''', BaseUrl, Vendor."No.");

            // Prepare JSON data
            JsonObject.Add('CompanyType', 'Vendor');
            JsonObject.Add('CompanyName', Vendor.Name);
            JsonObject.Add('CompanyNo', Vendor."No.");
            JsonObject.Add('CompanyLink', CompanyLink);

            // Add general details
            JsonObject.Add('CVR', Vendor."VAT Registration No.");
            JsonObject.Add('CountryCode', Vendor."Country/Region Code");
            JsonObject.Add('PhoneNumber', Vendor."Phone No.");
            JsonObject.Add('InvoiceEmail', Vendor."E-Mail");

            // Add location details
            JsonObject.Add('Address', Vendor.Address);
            JsonObject.Add('Address2', Vendor."Address 2");
            JsonObject.Add('City', Vendor.City);
            JsonObject.Add('PostalCode', Vendor."Post Code");
            JsonObject.Add('Country', Vendor."Country/Region Code");

            // Fetch main contact details
            if Contact.Get(Vendor."Primary Contact No.") then begin
                Clear(JsonContact);
                JsonContact.Add('ContactNo', Contact."No.");
                JsonContact.Add('Name', Contact.Name);
                JsonContact.Add('Phone', Contact."Phone No.");
                JsonContact.Add('MobilePhone', Contact."Mobile Phone No.");
                JsonContact.Add('DirectPhone', Contact."Phone No.");
                JsonContact.Add('Email', Contact."E-Mail");
                JsonContact.Add('JobTitle', Contact."Job Title");
                JsonObject.Add('PrimaryContact', JsonContact);
            end;
        end;

        // Convert JSON object to string
        JsonObject.WriteTo(JsonString);

        // Return the JSON response
        exit(JsonString);
    end;
}
