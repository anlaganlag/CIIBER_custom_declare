@echo off
echo Setting Git HTTP proxy to http://127.0.0.1:7890...
git config --global http.proxy http://127.0.0.1:7890

echo Setting Git HTTPS proxy to http://127.0.0.1:7890...
git config --global https.proxy http://127.0.0.1:7890

echo Verifying proxy settings:
echo HTTP Proxy: 
git config --global http.proxy
echo HTTPS Proxy: 
git config --global https.proxy

echo.
echo Proxy settings applied successfully!
echo You can now try your git push again.
echo.
echo To remove proxy settings later, run: unset_git_proxy.bat

pause