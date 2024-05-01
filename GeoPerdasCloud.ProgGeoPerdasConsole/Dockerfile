#See https://aka.ms/customizecontainer to learn how to customize your debug container and how Visual Studio uses this Dockerfile to build your images for faster debugging.

FROM mcr.microsoft.com/dotnet/runtime:6.0 AS base
WORKDIR /app

FROM mcr.microsoft.com/dotnet/sdk:6.0 AS build
WORKDIR /src
COPY ["GeoPerdasCloud.ProgGeoPerdasConsole/GeoPerdasCloud.ProgGeoPerdasConsole.csproj", "GeoPerdasCloud.ProgGeoPerdasConsole/"]
COPY ["GeoPerdasCloud.ProgGeoPerdas.Legacy/GeoPerdasCloud.ProgGeoPerdas.Legacy.csproj", "GeoPerdasCloud.ProgGeoPerdas.Legacy/"]
RUN dotnet restore "GeoPerdasCloud.ProgGeoPerdasConsole/GeoPerdasCloud.ProgGeoPerdasConsole.csproj"
COPY . .
WORKDIR "/src/GeoPerdasCloud.ProgGeoPerdasConsole"
RUN dotnet build "GeoPerdasCloud.ProgGeoPerdasConsole.csproj" -c Release -o /app/build

FROM build AS publish
RUN dotnet publish "GeoPerdasCloud.ProgGeoPerdasConsole.csproj" -c Release -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "GeoPerdasCloud.ProgGeoPerdasConsole.dll"]