pytest 実行手順

1.WinSCPで、mock_testフォルダはLinuxの適当な場所にコピーする (ここでは/tmp 直下とする)

2.カレントディレクトリを変更する。
# cd /tmp/mock_test

3.pytestを実行する。
# pytest -v --cov=src --cov-report term-missing test -x


