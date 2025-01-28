codeunit 60305 "NL Project Service"
{
    procedure GetProjectsForCompany(CompanyId: Code[20]; PageSize: Integer; PageNumber: Integer; IncludeClosedProjects: Boolean) returnValue: Text
    var
        Project: Record "Job";
        Projects: List of [Text];
        Result: Text[1024];
        StartIndex, EndIndex, Index : Integer;
    begin
        Project.SetRange("Bill-to Customer No.", CompanyId);
        if not IncludeClosedProjects then
            Project.SetRange(Status, Project.Status::Open);

        if Project.FindSet() then
            repeat
                Projects.Add(Format(Project.Description) + ' (' + Project."No." + ')');
            until Project.Next() = 0;

        StartIndex := ((PageNumber - 1) * PageSize) + 1;
        EndIndex := PageNumber * PageSize;
        if EndIndex > Projects.Count() then
            EndIndex := Projects.Count();

        for Index := StartIndex to EndIndex do
            Result += Projects.Get(Index) + ';';

        exit(Result.TrimEnd(';'));
    end;

    procedure GetTasksForProject(ProjectId: Code[20]) returnValue: Text
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

    procedure GetTaskAnalysis(TaskId: Code[20]) returnValue: Text
    var
        JsonObject: JsonObject;
        AnalysisObject: JsonObject;
        JsonString: Text;
    begin
        AnalysisObject.Add('TotalWorked', 46.25);
        AnalysisObject.Add('TotalAdjustments', -2.25);
        JsonObject.Add('Analysis', AnalysisObject);
        JsonObject.WriteTo(JsonString);

        exit(JsonString);
    end;
}
