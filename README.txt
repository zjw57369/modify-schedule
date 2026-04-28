============================================================
                    排班自动替换工具 v2.0
============================================================

【使用方式三选一】

方式1：exe版本（推荐，最简单）
  直接双击运行 modify_schedule.exe
  支持交互式菜单，可添加/删除规则

方式2：网页版本（推荐，界面美观）
  双击运行 "启动网页版.bat"
  然后在浏览器打开: http://localhost:5000
  支持拖拽上传、在线配置规则

方式3：Python脚本（适合开发者）
  命令行运行: python run_modify_schedule.py 文件名.xlsx

【文件说明】
  modify_schedule.exe   - Windows可执行程序
  app.py                - 网页版主程序
  启动网页版.bat         - 一键启动网页版
  templates/index.html  - 网页版界面
  requirements.txt      - Python依赖列表

【规则配置】
  所有规则保存在 schedule_rules.json
  在exe或网页版中添加的规则都会自动保存
