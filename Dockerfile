FROM microsoft/powershell:latest

ARG CLI_VERSION=latest

LABEL name="m365pnp/cli-microsoft365:${CLI_VERSION}" \
  description="Manage Microsoft 365 and SharePoint Framework projects on any platform" \
  homepage="https://pnp.github.io/cli-microsoft365" \
  maintainers="Waldek Mastykarz <waldek@mastykarz.nl>, \
  Velin Georgiev <velin.georgiev@gmail.com>, \
  Garry Trinder <garry.trinder@live.com>, \
  Albert-Jan Schot <appie@digiwijs.nl>, \
  Rabia Williams <rabiawilliams@gmail.com>, \
  Mark Powney <mpowney@icloud.com>, \
  Patrick Lamber <patrick@nubo.eu>"

ENV 0="/bin/bash" \
  SHELL="bash" \
  CLIMICROSOFT365_ENV="docker" \
  NPM_CONFIG_PREFIX=/home/cli-microsoft365/.npm-global \
  PATH=$PATH:/home/cli-microsoft365/.npm-global/bin

RUN apt-get -y update && apt-get install -y \
  curl \
  sudo \
  wget \
  bash-completion \
  && rm -rf /var/lib/apt/lists/*

RUN curl -sL https://deb.nodesource.com/setup_12.x | sudo -E bash - \
  && apt-get install -y nodejs

RUN useradd \
  --user-group \
  --system \
  --create-home \
  --no-log-init \
  cli-microsoft365
USER cli-microsoft365

WORKDIR /home/cli-microsoft365

RUN npm i -g @pnp/cli-microsoft365@${CLI_VERSION} --production

RUN bash -c 'm365 cli completion sh setup'
RUN pwsh -c 'm365 cli completion pwsh setup --profile $profile'

CMD [ "bash", "-l" ]