# 建置階段：使用 .NET 8 SDK 建置應用
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /src

# 複製 csproj 並還原 NuGet 套件
COPY *.csproj ./
RUN dotnet restore

# 複製所有檔案並建置
COPY . ./
RUN dotnet publish -c Release -o /app/publish

# 執行階段：使用 .NET 8 Runtime 執行應用
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /app
COPY --from=build /app/publish .

# 開放預設 Port
EXPOSE 80

# 設定啟動命令
ENTRYPOINT ["dotnet", "ShowPms.dll"]
