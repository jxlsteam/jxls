#!/usr/bin/env bash

# Deploy the generated Maven site to the JXLS SourceForge project website.
# Run this file with bash:
# bash ./deploy-site.sh
#
# Prerequisites:
#   1. Generate the site into ./target/site folder
#   2. Configure ~/.ssh/config with:
#
#        Host sourceforge-web
#            HostName web.sourceforge.net
#            User <your-sourceforge-username>
#
#   3. Make sure SSH authentication to SourceForge works:
#
#        ssh sourceforge-web
#
# The SourceForge web host provides a restricted file-transfer shell, so
# Maven's `site:deploy` is not used here. We use the native `scp` client.

scp -r ./target/site/. sourceforge-web:/home/project-web/jxls/htdocs/