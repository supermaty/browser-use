"""
Login helper module for Yuntu Data Scraper.

Provides functionality to check for existing login profiles and perform manual login.
"""
import asyncio
from pathlib import Path

from browser_use import Browser
from config import PROFILE_DIR, YUNTU_LOGIN_URL


async def check_profile_exists() -> bool:
    """
    Checks if the browser profile directory exists and has data.
    
    Returns:
        True if profile exists and has data, False otherwise.
    """
    if not PROFILE_DIR.exists():
        return False
    
    # Check if directory is not empty (has at least some files/folders)
    try:
        return any(PROFILE_DIR.iterdir())
    except Exception:
        return False


async def perform_manual_login() -> bool:
    """
    Launches a headful browser for the user to manually log in to Yuntu.
    
    The browser will stay open until:
    1. The user successfully logs in (detected by URL change)
    2. The user closes the browser
    3. A timeout occurs
    
    Returns:
        True if login was successful, False otherwise.
    """
    print("=" * 50)
    print("正在启动浏览器进行手动登录...")
    print(f"浏览器配置文件将保存到: {PROFILE_DIR}")
    print("=" * 50)
    
    browser = Browser(
        headless=False,
        user_data_dir=str(PROFILE_DIR),
    )
    
    login_successful = False
    
    try:
        # Start the browser session
        await browser.start()
        
        # Navigate to login page
        await browser.navigate_to(YUNTU_LOGIN_URL)
        
        print("\n" + "=" * 50)
        print("请在浏览器窗口中登录巨量云图")
        print("登录成功后，系统将自动检测并关闭浏览器")
        print("=" * 50 + "\n")
        
        # Poll for login success by checking URL
        max_wait_time = 300  # 5 minutes
        check_interval = 2  # Check every 2 seconds
        elapsed = 0
        
        while elapsed < max_wait_time:
            await asyncio.sleep(check_interval)
            elapsed += check_interval
            
            try:
                current_url = await browser.get_current_page_url()
                
                # Check if we've moved away from login page
                if current_url and "login" not in current_url.lower():
                    if "oceanengine.com" in current_url or "yuntu" in current_url.lower():
                        print(f"\n✅ 检测到登录成功！当前页面: {current_url}")
                        login_successful = True
                        
                        # Wait a bit more to ensure cookies are saved
                        await asyncio.sleep(3)
                        break
            except Exception as e:
                # Browser might have been closed by user
                print(f"检测状态时出错: {e}")
                break
        
        if not login_successful and elapsed >= max_wait_time:
            print("\n⚠️ 登录等待超时（5分钟），请重新尝试")
            
    except Exception as e:
        print(f"\n❌ 登录过程中出错: {e}")
    finally:
        try:
            await browser.stop()
        except Exception:
            pass
        print("\n浏览器已关闭。")
    
    return login_successful
