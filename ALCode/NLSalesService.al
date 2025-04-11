codeunit 60704 "NLSalesService"
{
    SingleInstance = true;
    Access = Public;

    procedure GetSalesQuotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        SalesHeader: Record "Sales Header";
        Quotes: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
        QuoteLink: Text;
    begin
        SalesHeader.SetRange("Sell-to Customer No.", CompanyId);
        SalesHeader.SetRange("Document Type", SalesHeader."Document Type"::Quote);

        if SalesHeader.FindSet() then
            repeat
                QuoteLink := BuildLink(41, SalesHeader."No.");
                Quotes.Add(Format(SalesHeader."Document Date") + ' (' + SalesHeader."No." + ') [' + QuoteLink + ']');
            until SalesHeader.Next() = 0;

        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        if (StartIndex + PageSize - 1) < Quotes.Count() then
            EndIndex := StartIndex + PageSize - 1
        else
            EndIndex := Quotes.Count();

        for Index := StartIndex to EndIndex do
            Result += Quotes.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetLastInvoices(CompanyId: Code[20]; Count: Integer) returnValue: Text
    var
        Invoice: Record "Sales Invoice Header";
        Invoices: List of [Text];
        Result: Text[1024];
        Index: Integer;
        InvoiceLink: Text;
    begin
        Invoice.SetRange("Sell-to Customer No.", CompanyId);
        Invoice.SetCurrentKey("Posting Date");
        Invoice.Ascending(false);

        if Invoice.FindSet() then
            repeat
                InvoiceLink := BuildLink(132, Invoice."No.");
                Invoices.Add(Format(Invoice."Posting Date") + ' (' + Invoice."No." + ') [' + InvoiceLink + ']');
            until (Invoice.Next() = 0) or (Invoices.Count() >= Count);

        for Index := 1 to Invoices.Count() do
            Result += Invoices.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetCreditNotes(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer) returnValue: Text
    var
        CreditMemo: Record "Sales Cr. Memo Entity Buffer";
        CreditNotes: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
        CreditNoteLink: Text;
    begin
        CreditMemo.SetRange("Sell-to Customer No.", CompanyId);

        if CreditMemo.FindSet() then
            repeat
                CreditNoteLink := BuildLink(44, CreditMemo."No.");
                CreditNotes.Add(
                    Format(CreditMemo."No.") + ': ' + Format(CreditMemo."Posting Date") +
                    ' [' + CreditNoteLink + ']'
                );
            until CreditMemo.Next() = 0;

        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        if (StartIndex + PageSize - 1) < CreditNotes.Count() then
            EndIndex := StartIndex + PageSize - 1
        else
            EndIndex := CreditNotes.Count();

        for Index := StartIndex to EndIndex do
            Result += CreditNotes.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetSalesQuoteLink(QuoteNo: Code[20]): Text
    begin
        exit(BuildLink(41, QuoteNo));
    end;

    procedure GetInvoiceLink(InvoiceNo: Code[20]): Text
    begin
        exit(BuildLink(132, InvoiceNo));
    end;

    procedure GetCreditNoteLink(CreditNoteNo: Code[20]): Text
    begin
        exit(BuildLink(44, CreditNoteNo));
    end;

    procedure BuildLink(PageId: Integer; RecordNo: Code[20]): Text
    begin
        exit(StrSubstNo('%1%2&page=%3&record=%4', GetBaseUrl(), GetCompanyQuery(), PageId, RecordNo));
    end;

    local procedure GetBaseUrl(): Text
    begin
        exit('https://nl-server.navilogic.dk/bc24-powerapp');
    end;

    local procedure GetCompanyQuery(): Text
    begin
        exit('/?company=CRONUS%20Danmark%20A%2fS');
    end;
}
