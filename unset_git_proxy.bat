@echo off
echo Removing Git HTTP proxy settings...
git config --global --unset http.proxy

echo Removing Git HTTPS proxy settings...
git config --global --unset https.proxy

echo Verifying proxy settings have been removed:
echo HTTP Proxy: 
git config --global http.proxy
echo HTTPS Proxy: 
git config --global https.proxy

echo.
echo Proxy settings removed successfully!
echo.
echo To set proxy settings again, run: set_git_proxy.bat

pause