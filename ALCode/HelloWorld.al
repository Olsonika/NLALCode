codeunit 60301 "GetCompaniesByEmail"
{
    [ServiceEnabled]
    procedure GetCompaniesForEmail(EmailAddress: Text[250]): List of [Text]
    var
        Customer: Record Customer;
        Contact: Record Contact;
        Companies: List of [Text];
        CompanyNameAndId: Text[250];
    begin
        // Search in Customer table
        if Customer.FindSet() then begin
            repeat
                if Customer."E-Mail" = EmailAddress then begin
                    CompanyNameAndId := Format(Customer.Name) + ' (' + Customer."No." + ')';
                    if not Companies.Contains(CompanyNameAndId) then
                        Companies.Add(CompanyNameAndId);
                end;
            until Customer.Next() = 0;
        end;

        // Search in Contact table
        if Contact.FindSet() then begin
            repeat
                if Contact."E-Mail" = EmailAddress then begin
                    // Look for the related customer
                    if Customer.Get(Contact."Company No.") then begin
                        CompanyNameAndId := Format(Customer.Name) + ' (' + Customer."No." + ')';
                        if not Companies.Contains(CompanyNameAndId) then
                            Companies.Add(CompanyNameAndId);
                    end;
                end;
            until Contact.Next() = 0;
        end;

        // Return the list of companies
        exit(Companies);
    end;
}
