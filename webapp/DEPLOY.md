# 部署到 kucun.riverline.com.cn（腾讯云服务器，122.51.255.195）

## 0. 前置：DNS 解析
去域名服务商（腾讯云DNSPod等）给 `riverline.com.cn` 加一条 A 记录：
```
主机记录: kucun
记录类型: A
记录值:   122.51.255.195
```

## 1. SSH 上服务器，拉代码
```bash
ssh <你的用户名>@122.51.255.195

sudo mkdir -p /opt/kucun
sudo chown $USER:$USER /opt/kucun
git clone https://github.com/mutoumoody-svg/kucun.git /opt/kucun
cd /opt/kucun/webapp
```

## 2. 装 Python 环境（如果服务器还没有 python3-venv）
```bash
sudo apt update && sudo apt install -y python3-venv python3-pip   # Debian/Ubuntu
# 或 CentOS: sudo yum install -y python3 python3-pip

python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## 3. 把已有的库存数据传上去（首次部署需要，不然"当月累计出库"等历史对比功能没数据可用）
在你自己电脑上（不是服务器）执行，把本地的 raw_data 和 output 文件夹整体传上去：
```bash
# 在本地 D:\Users\jingz\Documents\Claude项目\Inventory 目录下执行
scp -r raw_data <你的用户名>@122.51.255.195:/opt/kucun/
scp -r output   <你的用户名>@122.51.255.195:/opt/kucun/
```
（这两个文件夹是业务数据，没有放进git仓库，所以git clone之后要单独传一次。
以后新快照可以直接在网页上传，不需要再手动scp。）

## 4. 配置成系统服务，开机自启 + 常驻运行
```bash
sudo cp /opt/kucun/webapp/kucun-webapp.service /etc/systemd/system/
# 注意：service文件里 User=www-data，如果服务器上没有这个用户/权限不对，
# 改成你自己的用户名（跟ExecStart里venv路径权限要对得上）

sudo systemctl daemon-reload
sudo systemctl enable kucun-webapp
sudo systemctl start kucun-webapp
sudo systemctl status kucun-webapp   # 确认是 active (running)
```

## 5. 配置 nginx 反向代理
```bash
sudo cp /opt/kucun/webapp/nginx-kucun.conf /etc/nginx/conf.d/kucun.conf
sudo nginx -t        # 测试配置没问题
sudo systemctl reload nginx
```

## 6. 验证
浏览器打开 `http://kucun.riverline.com.cn`，应该能看到上传页面。

## 7.（可选）配 HTTPS
如果服务器上有 certbot：
```bash
sudo certbot --nginx -d kucun.riverline.com.cn
```

## 以后更新代码怎么办
本地改完代码、git push 之后，登录服务器执行：
```bash
cd /opt/kucun
git pull
sudo systemctl restart kucun-webapp
```
