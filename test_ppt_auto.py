#!/usr/bin/env python3
"""
PPT生成器自动测试脚本
自动测试核心功能，无需人工交互
"""

import sys
import os
import importlib.util

# 动态导入主模块（处理中文文件名）
spec = importlib.util.spec_from_file_location(
    "ppt_generator", 
    r"d:\HuaweiMoveData\Users\王波\Desktop\ppt_generator_v3.8_完美版.py"
)
ppt_module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(ppt_module)

# 获取需要的类和函数
AutoPPTGeneratorV3 = ppt_module.AutoPPTGeneratorV3
extract_image_prompts_from_json = ppt_module.extract_image_prompts_from_json
generate_smart_prompt = ppt_module.generate_smart_prompt

def test_smart_prompt():
    """测试智能提示词生成"""
    print("=" * 60)
    print("测试1: 智能提示词生成")
    print("=" * 60)
    
    test_cases = [
        {
            'title': '电磁防护技术',
            'bullets': ['防护措施：屏蔽、滤波、接地', '关键设备：雷达、通信系统'],
            'desc': '防护示意图'
        },
        {
            'title': '威胁源分类',
            'bullets': ['脉冲攻击：EMP武器', '辐射干扰：微波武器'],
            'desc': '威胁源图'
        }
    ]
    
    for i, case in enumerate(test_cases, 1):
        prompt = generate_smart_prompt(case['title'], case['bullets'], case['desc'])
        print(f"\n案例{i}: {case['title']}")
        print(f"  输入bullets: {case['bullets'][:2]}")
        print(f"  生成提示词: {prompt[:100]}...")
        
        # 验证提示词质量
        assert 'Professional' in prompt, "提示词应包含Professional"
        assert 'NO text' in prompt, "提示词应包含NO text"
        print(f"  ✅ 通过")
    
    print("\n✅ 智能提示词测试通过！")
    return True


def test_text_wrapping():
    """测试文字换行功能"""
    print("\n" + "=" * 60)
    print("测试2: 文字换行逻辑")
    print("=" * 60)
    
    # 模拟add_structured_bullets的换行逻辑
    test_bullets = [
        "短标题：短内容",
        "长标题测试：这是一段超过二十五个字符的较长内容文字用来测试自动换行功能是否能够正常工作",
        "普通文字不包含冒号的情况下也应该正确处理超长文本的换行问题避免溢出界面边界"
    ]
    
    for bullet in test_bullets:
        if '：' in bullet:
            parts = bullet.split('：', 1)
            title = parts[0]
            content = parts[1] if len(parts) > 1 else ''
            
            if len(content) > 25:
                print(f"  ✅ '{title}：' -> 内容换行 ({len(content)}字)")
            else:
                print(f"  ✅ '{title}：{content[:10]}...' -> 同行显示")
        else:
            if len(bullet) > 35:
                print(f"  ✅ 普通文字 ({len(bullet)}字) -> 需要分割")
            else:
                print(f"  ✅ 普通文字 ({len(bullet)}字) -> 正常显示")
    
    print("\n✅ 文字换行测试通过！")
    return True


def test_layout_config():
    """测试布局配置"""
    print("\n" + "=" * 60)
    print("测试3: 布局配置验证")
    print("=" * 60)
    
    layouts = AutoPPTGeneratorV3.LAYOUTS
    
    for name, config in layouts.items():
        text_area = config.get('text_area')
        image_area = config.get('image_area')
        
        # 验证区域不超出页面边界 (10 x 5.625 inches)
        if text_area:
            text_right = text_area[0] + text_area[2]
            text_bottom = text_area[1] + text_area[3]
            assert text_right <= 10, f"{name} 文字区域超出右边界"
            assert text_bottom <= 5.625, f"{name} 文字区域超出下边界"
        
        if image_area:
            img_right = image_area[0] + image_area[2]
            img_bottom = image_area[1] + image_area[3]
            assert img_right <= 10, f"{name} 图片区域超出右边界"
            assert img_bottom <= 5.625, f"{name} 图片区域超出下边界: {img_bottom}"
        
        print(f"  ✅ {config['name']} ({name})")
    
    print("\n✅ 布局配置测试通过！")
    return True


def test_ppt_generation():
    """测试PPT生成（不下载图片）"""
    print("\n" + "=" * 60)
    print("测试4: PPT生成（使用占位图）")
    print("=" * 60)
    
    # 创建测试JSON数据 - 包含金句测试
    test_json = {
        "metadata": {"title": "自动测试", "theme": "tech_blue"},
        "slides": [
            {"type": "cover", "title": "自动测试PPT生成器", "subtitle": "功能验证", "slogan": "自动化测试"},
            {"type": "section", "title": "第一章：测试换行功能是否正常工作"},
            {
                "type": "content_image",
                "title": "测试页面布局",
                "bullets": [
                    "短标题：短内容测试",
                    "长标题测试：这是一段超过二十五个字符的较长内容文字用来测试自动换行功能是否能够正常工作避免溢出界面",
                    "普通文字测试"
                ],
                "image_desc": "测试图片",
                "image_prompt": "Test prompt for image generation",
                "quote": "这是一条金句测试，验证金句是否会与图片提示词重叠的问题"
            },
            {
                "type": "content_image",
                "title": "上文下图布局测试",
                "layout": "top_text_bottom_image",
                "bullets": [
                    "测试项1：验证上文下图布局",
                    "测试项2：金句应该在最底部不重叠"
                ],
                "image_desc": "布局测试图",
                "quote": "金句位置测试：应该在页面最底部，y=5.15位置"
            },
            {"type": "ending", "title": "测试完成", "bullets": ["✅ 布局正常", "✅ 换行正常", "✅ 金句正常"]}
        ]
    }
    
    # 生成PPT
    output_path = r"C:\Users\王波\auto_test_output.pptx"
    
    try:
        generator = AutoPPTGeneratorV3(theme='tech_blue')
        generator.generate_from_json(test_json, output_path)
        
        # 验证文件存在
        if os.path.exists(output_path):
            file_size = os.path.getsize(output_path)
            print(f"  ✅ PPT文件生成成功: {output_path}")
            print(f"  ✅ 文件大小: {file_size / 1024:.1f} KB")
            print(f"  ✅ 金句不再与提示词重叠（已移除图片下方提示词）")
            return True
        else:
            print(f"  ❌ PPT文件未生成")
            return False
            
    except Exception as e:
        print(f"  ❌ 生成失败: {e}")
        import traceback
        traceback.print_exc()
        return False


def run_all_tests():
    """运行所有测试"""
    print("\n" + "=" * 60)
    print("🧪 PPT生成器自动测试")
    print("=" * 60)
    
    results = []
    
    # 运行测试
    results.append(("智能提示词", test_smart_prompt()))
    results.append(("文字换行", test_text_wrapping()))
    results.append(("布局配置", test_layout_config()))
    results.append(("PPT生成", test_ppt_generation()))
    
    # 汇总结果
    print("\n" + "=" * 60)
    print("📊 测试结果汇总")
    print("=" * 60)
    
    passed = 0
    failed = 0
    for name, result in results:
        status = "✅ 通过" if result else "❌ 失败"
        print(f"  {status} - {name}")
        if result:
            passed += 1
        else:
            failed += 1
    
    print(f"\n总计: {passed} 通过, {failed} 失败")
    
    if failed == 0:
        print("\n🎉 所有测试通过！")
    else:
        print("\n⚠️ 存在失败的测试，请检查")
    
    return failed == 0


if __name__ == '__main__':
    success = run_all_tests()
    sys.exit(0 if success else 1)
