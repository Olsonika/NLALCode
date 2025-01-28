codeunit 60304 "NL Invoice Service"
{
    procedure GetSalesQuotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        SalesHeader: Record "Sales Header";
        Quotes: List of [Text];
        Result: Text[1024];
        StartIndex, EndIndex, Index : Integer;
    begin
        SalesHeader.SetRange("Sell-to Customer No.", CompanyId);
        SalesHeader.SetRange("Document Type", SalesHeader."Document Type"::Quote);

        if SalesHeader.FindSet() then
            repeat
                Quotes.Add(Format(SalesHeader."No.") + ': ' + Format(SalesHeader."Document Date"));
            until SalesHeader.Next() = 0;

        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        EndIndex := PageNumber * PageSize;
        if EndIndex > Quotes.Count() then
            EndIndex := Quotes.Count();

        for Index := StartIndex to EndIndex do
            Result += Quotes.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetLastInvoices(CompanyId: Code[20]; Count: Integer) returnValue: Text
    var
        Invoice: Record "Sales Invoice Header";
        Result: Text[1024];
        Index: Integer;
    begin
        Invoice.SetRange("Sell-to Customer No.", CompanyId);
        Invoice.SetCurrentKey("Posting Date");
        Invoice.Ascending(false);

        if Invoice.FindSet() then
            repeat
                Result += Format(Invoice."No.") + ': ' + Format(Invoice."Posting Date") + ';';
            until (Invoice.Next() = 0) or (Index >= Count);

        exit(Result.TrimEnd(';'));
    end;

    procedure GetCreditNotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        CreditMemo: Record "Sales Cr. Memo Entity Buffer";
        Result: Text[1024];
        CreditNotes: List of [Text];
        StartIndex, EndIndex, Index : Integer;
    begin
        CreditMemo.SetRange("Sell-to Customer No.", CompanyId);
        if CreditMemo.FindSet() then
            repeat
                CreditNotes.Add(Format(CreditMemo."No.") + ': ' + Format(CreditMemo."Posting Date"));
            until CreditMemo.Next() = 0;

        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        EndIndex := PageNumber * PageSize;
        if EndIndex > CreditNotes.Count() then
            EndIndex := CreditNotes.Count();

        for Index := StartIndex to EndIndex do
            Result += CreditNotes.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;
}
