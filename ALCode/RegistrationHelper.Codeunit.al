codeunit 60302 "NL Registration Helper"
{
    procedure SelectCustomerTemplate(LanguageCode: Text; CountryCode: Text): Code[20]
    var
        CountryRegion: Record "Country/Region";
        CustomerTemplate: Record "Customer Templ.";
        TemplateCode: Code[20];
    begin
        if not CountryRegion.Get(CountryCode) then
            Error('Country/Region Code %1 not found.', CountryCode);

        if CountryCode = 'DK' then
            TemplateCode := 'DK'
        else
            if CountryRegion."EU Country/Region Code" <> '' then
                TemplateCode := 'EU'
            else
                TemplateCode := 'NON-EU';

        if CustomerTemplate.Get(TemplateCode) then
            exit(TemplateCode)
        else
            // If the specific template is not found, fallback to language code
            if CustomerTemplate.FindSet() then
                repeat
                    if CustomerTemplate."Language Code" = LanguageCode then begin
                        TemplateCode := CustomerTemplate.Code;
                        exit(TemplateCode);
                    end;
                until CustomerTemplate.Next() = 0;
        exit('');
    end;

    procedure NewContact(
        ContactNoForCompany: Code[20];
        FirstAndLastName: Text;
        DirectPhoneNumber: Text;
        MobilePhoneNumber: Text;
        Email: Text[80];
        Title: Text
    ) result: Code[20]
    var
        Contact: Record Contact;
        MarketingSetup: Record "Marketing Setup"; // Added Marketing Setup
        NoSeries: Codeunit "No. Series";
        ContactNoForContact: Code[20];
        ContactNoSeriesCode: Text[20]; // Added for contact number series
    begin
        // Retrieve Contact Nos. from Marketing Setup
        if MarketingSetup.Get() then begin
            ContactNoSeriesCode := MarketingSetup."Contact Nos.";
            if ContactNoSeriesCode = '' then
                Error('No. Series for Contact Nos. is not defined in Marketing Setup.');
            ContactNoForContact := NoSeries.GetNextNo(ContactNoSeriesCode, Today(), true);
        end else
            Error('Marketing Setup not found.');

        Contact.Init();
        Contact.Type := Contact.Type::Person;
        Contact."No." := ContactNoForContact;

        Contact.Validate("Company No.", ContactNoForCompany);
        Contact.Validate(Name, FirstAndLastName);
        Contact.Validate("Phone No.", DirectPhoneNumber);
        Contact.Validate("Mobile Phone No.", MobilePhoneNumber);
        ValidateEmail(Email);
        Contact."E-Mail" := Email;
        Contact.Validate("Job Title", Title);

        if Contact.Insert() then
            result := ContactNoForContact;
    end;

    procedure FilterOutLetters(Input: Text): Text
    var
        i: Integer;
        ResultText: Text;
        Char: Text[1];
    begin
        for i := 1 to StrLen(Input) do begin
            Char := CopyStr(Input, i, 1);
            if (Char >= '0') and (Char <= '9') or (Char = '+') or (Char = '-') or (Char = ' ') then
                ResultText := ResultText + Char;
        end;
        exit(ResultText);
    end;

    procedure ValidateEmail(email: text)
    var
        Regex: Codeunit Regex;
        InvalidEmail_Err: Label 'Invalid Email Address';
        Pattern: Text;
    begin
        Pattern := '^[\w!#$%&*+\-/=?\^_`{|}~]+(\.[\w!#$%&*+\-/=?\^_`{|}~]+)*@((([\-\w]+\.)+[a-zA-Z]{2,4})|(([0-9]{1,3}\.){3}[0-9]{1,3}))$';

        if not Regex.IsMatch(email, Pattern, 0)
        then
            Error(InvalidEmail_Err);
    end;
}
