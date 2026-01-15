#!/bin/bash

# 1. Create tarball
echo "Creating trace.tar.gz..."
tar -zcvf trace.tar.gz assets manifest.xml trace.html trace.js multidim_dag_resolution

# 2. Send to server
echo "Sending trace.tar.gz to server..."
scp trace.tar.gz root@aliyun.ecs.us.west:/var/www/excel-add-in

# 3. Remove local tarball
echo "Cleaning up local trace.tar.gz..."
rm trace.tar.gz

# 4. Execute remote script
echo "Executing remote deploy script..."
ssh root@aliyun.ecs.us.west "/var/www/excel-add-in/deploy.sh"

echo "Deployment finished."
