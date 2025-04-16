codeunit 60705 "NLProjectService"
{
    SingleInstance = true;
    Access = Public;

    procedure GetProjectsForCompany(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer; IncludeClosedProjects: Boolean): Text
    var
        Project: Record "Job";
        Companies: Record Customer;
        Projects: List of [Text];
        Result: Text[1024];
        StartIndex: Integer;
        EndIndex: Integer;
        Index: Integer;
        BaseUrl: Text;
        ProjectLink: Text;
    begin
        BaseUrl := 'https://nl-server.navilogic.dk/bc24-intern';

        if not Companies.Get(CompanyId) then
            Error('Company with ID "%1" not found.', CompanyId);

        Project.SetRange("Bill-to Customer No.", Companies."No.");
        if not IncludeClosedProjects then
            Project.SetRange(Status, Project.Status::Open);

        if Project.FindSet() then
            repeat
                ProjectLink := StrSubstNo('%1?page=88&filter=''No.''%%20IS%%20''%2''', BaseUrl, Project."No.");
                Projects.Add(Format(Project.Description) + ' (' + Project."No." + ') [' + ProjectLink + ']');
            until Project.Next() = 0;

        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        if StartIndex + PageSize - 1 < Projects.Count() then
            EndIndex := StartIndex + PageSize - 1
        else
            EndIndex := Projects.Count();

        for Index := StartIndex to EndIndex do
            Result += Projects.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetTasksForProject(ProjectId: Code[20]): Text
    var
        JobTask: Record "Job Task";
        Result: Text[1024];
    begin
        JobTask.SetRange("Job No.", ProjectId);
        if not JobTask.FindFirst() then
            Error('Project with ID "%1" not found.', ProjectId);

        if JobTask.FindSet() then
            repeat
                Result += Format(JobTask."Job Task No.") + ', ' + JobTask.Description + ';';
            until JobTask.Next() = 0;

        exit(Result.TrimEnd(';'));
    end;

    procedure GetTaskAnalysis(TaskId: Code[20]): Text
    var
        JsonObject: JsonObject;
        AnalysisObject: JsonObject;
        ChargeableObject: JsonObject;
        FreeOfChargeObject: JsonObject;
        JsonString: Text;
    begin
        AnalysisObject.Add('TotalWorked', 46.25);
        AnalysisObject.Add('TotalAdjustments', -2.25);
        AnalysisObject.Add('TotalFreeOfCharge', -2.00);
        AnalysisObject.Add('TotalChargeable', 42.00);
        AnalysisObject.Add('Total100PercentDiscount', 0.50);
        AnalysisObject.Add('TotalShownOnInvoice', 42.50);
        AnalysisObject.Add('ExpectedBilling', 0.50);

        ChargeableObject.Add('Invoiced', 41.50);
        ChargeableObject.Add('Registered', 0.50);
        ChargeableObject.Add('TotalChargeable', 42.00);

        FreeOfChargeObject.Add('Internal', 1.50);
        FreeOfChargeObject.Add('100PercentDiscountInvoiced', 0.50);
        FreeOfChargeObject.Add('TotalFreeOfCharge', 2.00);

        JsonObject.Add('Analysis', AnalysisObject);
        JsonObject.Add('Chargeable', ChargeableObject);
        JsonObject.Add('FreeOfCharge', FreeOfChargeObject);

        JsonObject.WriteTo(JsonString);

        exit(JsonString);
    end;
}
