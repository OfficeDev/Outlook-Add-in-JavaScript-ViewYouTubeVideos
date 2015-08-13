#!/usr/bin/env bash

# User runtime configuration
PORT=8443;

# Environment Variables
MANIFEST=manifest.xml
PROTOCOL=https:\\/\\/;
HOST=0.0.0.0;
FP="$PROTOCOL$HOST:$PORT";

# Delims
T_HOST='<?bash $ip ?>';
T_PORT='<?bash $port ?>';

# Cleans any previously generated manifests
function clean {
  echo "Cleaning project files"
  [[ -f "$MANIFEST" ]] && rm $MANIFEST;
}

# Install dependencies
function exec_bundler {
  echo "Checking dependencies"
  if hash bundler 2>/dev/null; then
      echo "Bundler already installed. Hooray."
  else
      echo "Bundler not installed - installing requires sudoer status."
      echo "Enter password to proceed with Bundler installation"
      sudo gem install bundler;
  fi
  echo "Verifying gems."
  bundle install;
}

# Generates an app manifest for this machine
function generate_manifest {
  echo "Generating add-in manifest"
  sed "s/$T_HOST/$FP/g" .template/manifest/manifest.xml > manifest.xml;
}

# Generates a Sinatra script listening on the same port for which the manifest
# was generated
function generate_sinatra {
  echo "Generating server script"
  sed "s/$T_PORT/$PORT/g" .template/serv/app.rb > app.rb
}

function inform_cert {
  # Check if they have a cert - do they?
  # If they don't tell them to run the script...
  if [ ! -f cert/server.crt ] && [ ! -f cert/server.key ]; then
      echo "No certificate installed. Generating..";
      cd cert;
      ./ss_certgen.sh;
  fi
}

## Main
clean
sleep 1
exec_bundler
sleep 1
generate_manifest
sleep 1
generate_sinatra
sleep 1
inform_cert;
