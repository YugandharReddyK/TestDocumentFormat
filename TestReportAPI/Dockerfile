FROM mcr.microsoft.com/dotnet/framework/sdk:4.8 AS build-env

WORKDIR /app
COPY . /app
RUN nuget.exe restore .\TestDocumentFormat.sln -SolutionDirectory ../ -Verbosity normal
RUN MSBuild.exe .\TestReportAPI\TestReportAPI.csproj /t:build /p:Configuration=Release /p:OutputPath=/app/out

FROM mcr.microsoft.com/dotnet/framework/runtime:4.8 AS runtime
WORKDIR /app
COPY --from=build-env /app/out .
ENTRYPOINT ["TestReportAPI.exe"]