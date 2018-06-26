#!/bin/bash
# author:senman
# url:www.monqin.com

echo "====================================================================";
echo "您好，我是senman，不一样的烟火就是我，现在为您部署openoffice环境";
echo "====================================================================";

#判断系统位数
s_bit=`getconf LONG_BIT`;
echo "您的系统是$s_bit位 " ;

JAVA_VERSION=`java -version`;
if [ $? -ne 0 ]; then
   echo "请先安装java环境，并配置好jdk";
   exit;
else
   echo "jdk 已经安装";
fi

mkdir -p  /home/software/download/;
chmod -R 777 /home/software/download/;

#防止入坑1
echo "开始安装x-wimdows-system";
yum groupinstall "X Window System";

echo "开始创建office的下载目录,/home/software/download/office/ ";

mkdir -p  /home/software/download/openoffice/;
echo "给目录赋予权限 /home/office/ ";
chmod -R 777 /home/software/download/openoffice/;
echo "开始下载 ";
cd  /home/software/download/openoffice/;

#根据系统位数选择不同的安装包
if [ $s_bit==64 ];then
   echo "开始下载64位的openoffice"; 
   wget -vcO download https://sourceforge.net/projects/openofficeorg.mirror/files/4.1.5/binaries/zh-CN/Apache_OpenOffice_4.1.5_Linux_x86-64_install-rpm_zh-CN.tar.gz/download; 
else
   echo "开始下载32位的openoffice"；
   wget -vcO  download https://sourceforge.net/projects/openofficeorg.mirror/files/4.1.5/binaries/zh-CN/Apache_OpenOffice_4.1.5_Linux_x86_install-rpm_zh-CN.tar.gz/download;
fi


#判断文件是否下载成功
if  [ -f "/home/software/download/openoffice/download" ];then
     echo "下载完毕，开始解压文件";
else
     echo "文件下载失败";
     exit
fi

cd  /home/software/download/openoffice/;
tar -xvzf download;

cd  /home/software/download/openoffice/;
cd   zh-CN/RPMS/;
echo "开始安装openoffice";
rpm -ivh *.rpm;


cd desktop-integration/;
rpm -ivh openoffice4.1.5-redhat-menus-4.1.5-9789.noarch.rpm;

#防止入坑2；
yum install libXext.x86_64;
cp -a /usr/lib64/libXext.so.6 /opt/openoffice4/program/;


#防止入坑3
yum install freetype;
cp -a /usr/lib64/libfreetype.so.6 /opt/openoffice4/program/;

echo "安装完成,若是转换出现中文乱码则请注意 jre、 openoffice、系统的字体是否有中文字体";


#启动命令
echo "正在启动openoffice服务………………………………";
sudo /opt/openoffice4/program/soffice -headless -accept="socket,host=127.0.0.1,port=8100;urp;" -nofirststartwizard &

netstat -lnp |grep 8100;

