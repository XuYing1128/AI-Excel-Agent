# -*- coding: utf-8 -*-
"""socket 诊断脚本：排查 WinError 10014 是否为安全软件导致。

跑这个脚本后，按提示用浏览器访问 http://127.0.0.1:9999
- 如果能看到 "OK" → socket 正常，问题在 streamlit 的复杂交互
- 如果报 10014 / 连不上 → 安全软件彻底 hook 了网络栈，必须关安全软件或加白名单
"""
import socket
import sys

def main():
    print("=" * 50)
    print("socket 诊断 (排查 WinError 10014)")
    print("=" * 50)

    s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
    try:
        s.bind(("127.0.0.1", 9999))
        s.listen(1)
    except OSError as e:
        print(f"[失败] 无法监听 9999 端口: {e}")
        print("→ 可能端口被占用或权限不足。")
        input("按回车退出...")
        return

    print("[1/3] 监听 127.0.0.1:9999 成功")
    print("[2/3] 现在请用浏览器访问: http://127.0.0.1:9999")
    print("      (或另开一个 cmd 跑: curl http://127.0.0.1:9999)")
    print("[3/3] 等待连接中... (最多等 60 秒)")
    print("-" * 50)

    s.settimeout(60)
    try:
        conn, addr = s.accept()
        print(f"[accept 成功] 对端地址: {addr}")
        peer = conn.getpeername()
        print(f"[getpeername 成功] {peer}")
        conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Length: 2\r\n\r\nOK")
        conn.close()
        print()
        print("=" * 50)
        print("✅ socket 完全正常！能用 getpeername。")
        print("→ 说明 10014 是 streamlit 的某个交互触发，不是安全软件。")
        print("=" * 50)
    except socket.timeout:
        print("[超时] 60 秒内没有连接进来。请确认你访问了 http://127.0.0.1:9999")
    except OSError as e:
        print()
        print("=" * 50)
        print(f"❌ 报错了: {e}")
        if "10014" in str(e):
            print("→ 确认是 WinError 10014！这是安全软件(360/火绒/校园管控) hook 了网络栈。")
            print("→ 解决: 关闭安全软件/防火墙, 或把本程序加入白名单, 或换台电脑试。")
        print("=" * 50)
    finally:
        s.close()

    input("\n按回车退出...")


if __name__ == "__main__":
    main()
