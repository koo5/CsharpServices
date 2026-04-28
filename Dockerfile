FROM mcr.microsoft.com/dotnet/sdk:8.0@sha256:ec47373b40fc22a31c891c5be8df83f8cbadc24e785b10899545e33a1a1c2183 AS build-env
WORKDIR /App

COPY . ./
RUN dotnet restore

RUN dotnet publish -c Release -o out



# Build runtime image
FROM mcr.microsoft.com/dotnet/aspnet:8.0@sha256:8c7de8ca26baa56530cdc3d55b5563420ee2a302619ad808c53ec64613143771

RUN mkdir /tmp; chmod 1777 /tmp
RUN apt update; apt install -y curl fonts-noto-core


WORKDIR /App
COPY --from=build-env /App/out .

ARG uid=1000
ARG gid=1000
RUN groupadd -g $gid myuser && useradd -lm -u $uid -g $gid myuser

VOLUME /app/server_root/tmp
USER myuser

VOLUME /App/RdfTemplates.n3
VOLUME /home/myuser/.dotnet

ENV COMPlus_EnableDiagnostics=0

ENTRYPOINT ["dotnet", "CsharpServices.dll"]
ENV HC="curl --noproxy '*' -L -S --fail http://127.0.0.1:17789/health"









#start interval is the time between health checks during the start period. This option requires Docker Engine version 25.0 or later.
USER root
RUN mkdir /healthcheck
RUN chown myuser:myuser /healthcheck
USER myuser
HEALTHCHECK --interval=2s --timeout=600s CMD { test -f /healthcheck/ok && echo -n "sleep.." && sleep 555 && echo -n "hc..>> " && /bin/sh -c '/bin/sh -c "$HC"' && echo -n " <<..hc passed." ; } || { echo -n "hc..>> " && echo -n "$HC" && echo -n " <<... " && /bin/sh -c '/bin/sh -c "$HC"' && echo -n " << first hc passed." && touch /healthcheck/ok && echo " flag stored." ; } 


