# Install
1. clone仓库
```bash
git clone https://github.com/eiclpy/sendemail.git
cd sendemail
```
2. 安装python3及pip3
```bash
sudo apt-get install python3 python3-pip
```
3. 安装依赖
```bash
pip3 install -r requirements.txt --user
```
4. 运行
```bash
python3 main.py
```
# Configure
修改config.py中listen及port的值，默认端口5000
修改config.py中的username及password