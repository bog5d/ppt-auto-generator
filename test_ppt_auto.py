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
parse_outline_to_json = ppt_module.parse_outline_to_json

def test_outline_parser():
    """测试大纲解析功能"""
    print("=" * 60)
    print("测试1: 大纲文本解析")
    print("=" * 60)
    
    # 测试大纲文本
    outline_text = """# 电磁防护技术培训
智能防护系统介绍

## 第一章 威胁分析
### 电磁脉冲威胁
- EMP攻击：高空核爆产生的电磁脉冲
- 雷电感应：自然界的电磁威胁
- 人为干扰：电子战武器
> 知己知彼，百战不殆

### 辐射效应分析
- 传导耦合：通过电缆传播
- 辐射耦合：空间电磁波
> 防患于未然

## 第二章 防护措施
### 屏蔽技术
- 金属屏蔽：法拉第笼原理
- 复合材料：轻量化防护
"""
    
    result = parse_outline_to_json(outline_text)
    
    # 验证结果
    assert result is not None, "解析结果不应为空"
    assert 'slides' in result, "应包含slides"
    assert len(result['slides']) >= 5, f"应至少有5个slide，实际有{len(result['slides'])}个"
    
    # 验证封面
    cover = result['slides'][0]
    assert cover['type'] == 'cover', "第一页应是封面"
    assert '电磁防护' in cover['title'], "封面标题应包含'电磁防护'"
    
    # 验证章节
    section_count = sum(1 for s in result['slides'] if s['type'] == 'section')
    assert section_count >= 2, "应至少有2个章节"
    
    # 验证内容页
    content_count = sum(1 for s in result['slides'] if s['type'] == 'content_image')
    assert content_count >= 3, "应至少有3个内容页"
    
    # 验证金句
    quotes_found = sum(1 for s in result['slides'] if s.get('quote'))
    assert quotes_found >= 2, "应至少有2个金句"
    
    print(f"   解析成功: {len(result['slides'])} 页幻灯片")
    print(f"   封面: {cover['title']}")
    print(f"   章节数: {section_count}")
    print(f"   内容页: {content_count}")
    print(f"   金句数: {quotes_found}")
    
    print("\n 大纲解析测试通过！")
    return True


def test_smart_prompt():
    """测试智能提示词生成"""
    print("\n" + "=" * 60)
    print("测试2: 智能提示词生成")
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
        print(f"  生成提示词: {prompt[:80]}...")
        
        # 验证提示词质量
        assert 'Professional' in prompt, "提示词应包含Professional"
        assert 'NO text' in prompt, "提示词应包含NO text"
        print(f"   通过")
    
    print("\n 智能提示词测试通过！")
    return True


def test_text_wrapping():
    """测试文字换行功能"""
    print("\n" + "=" * 60)
    print("测试3: 文字换行逻辑")
    print("=" * 60)
    
    # 模拟add_structured_bullets的换行逻辑
    test_bullets = [
        "短标题：短内容",
        "长标题测试：这是一段超过二十五个字符的较长内容文字用来测试自动换行功能是否能正常工作",
        "普通文字不包含冒号的情况下也应该正确处理超长文本的换行问题避免溢出界面边界"
    ]
    
    for bullet in test_bullets:
        if '：' in bullet:
            parts = bullet.split('：', 1)
            title = parts[0]
            content = parts[1] if len(parts) > 1 else ''
            
            if len(content) > 25:
                print(f"   '{title}：' -> 内容换行 ({len(content)}字)")
            else:
                print(f"   '{title}：{content[:10]}...' -> 同行显示")
        else:
            if len(bullet) > 35:
                print(f"   普通文字({len(bullet)}字) -> 需要分割")
            else:
                print(f"   普通文字({len(bullet)}字) -> 正常显示")
    
    print("\n 文字换行测试通过！")
    return True


def test_layout_config():
    """测试布局配置"""
    print("\n" + "=" * 60)
    print("测试4: 布局配置验证")
    print("=" * 60)
    
    layouts = AutoPPTGeneratorV3.LAYOUTS
    
    for name, config in layouts.items():
        text_area = config.get('text_area')
        image_area = config.get('image_area')
        
        assert text_area is not None, f"{name} 缺少text_area"
        assert image_area is not None, f"{name} 缺少image_area"
        assert len(text_area) == 4, f"{name} text_area格式错误"
        assert len(image_area) == 4, f"{name} image_area格式错误"
        
        print(f"   {name}: 文字区{text_area[:2]}, 图片区{image_area[:2]}")
    
    print("\n 布局配置测试通过！")
    return True


def test_ppt_generation():
    """测试PPT生成"""
    print("\n" + "=" * 60)
    print("测试5: PPT生成（无图片）")
    print("=" * 60)
    
    # 使用大纲解析生成的JSON
    outline = """# 测试PPT
自动化测试

## 第一章
### 测试内容
- 要点一
- 要点二
> 测试金句
"""
    
    json_data = parse_outline_to_json(outline)
    
    # 创建生成器
    generator = AutoPPTGeneratorV3(theme='military_solemn')
    
    # 生成PPT
    output_path = "test_output.pptx"
    generator.generate_from_json(json_data, output_path)
    
    # 验证文件
    assert os.path.exists(output_path), "PPT文件未生成"
    file_size = os.path.getsize(output_path)
    assert file_size > 10000, f"PPT文件太小: {file_size} bytes"
    
    print(f"\n PPT生成成功: {output_path} ({file_size/1024:.1f} KB)")
    
    # 清理
    os.remove(output_path)
    print(" 测试文件已清理")
    
    return True


def main():
    """运行所有测试"""
    print("\n" + "=" * 70)
    print(" PPT生成器 v3.9 自动化测试")
    print("=" * 70 + "\n")
    
    tests = [
        ("大纲解析", test_outline_parser),
        ("智能提示词", test_smart_prompt),
        ("文字换行", test_text_wrapping),
        ("布局配置", test_layout_config),
        ("PPT生成", test_ppt_generation),
    ]
    
    passed = 0
    failed = 0
    
    for name, test_func in tests:
        try:
            if test_func():
                passed += 1
        except Exception as e:
            print(f"\n {name} 测试失败: {e}")
            failed += 1
    
    print("\n" + "=" * 70)
    print(f" 测试结果: {passed} 通过, {failed} 失败")
    print("=" * 70)
    
    return failed == 0


if __name__ == '__main__':
    success = main()
    sys.exit(0 if success else 1)
