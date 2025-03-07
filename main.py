import os
import time
import cv2

from icetools.window_tools import fast_capture_full
from icetools.operation_tools import OperationTool as OPT

from OmnUtil.utils import check_ocr_box, get_yolo_model, get_caption_model_processor, get_som_labeled_img
import torch
import numpy as np
from PIL import Image
import pyautogui
import pyperclip
from skimage.metrics import structural_similarity as ssim

DEVICE = torch.device('cuda' if torch.cuda.is_available() else 'cpu')
SCREEN_WIDTH, SCREEN_HEIGHT = pyautogui.size()

# 模板匹配使用到的模板
from template_loader import *

# 调用api（可选）
# from openai import OpenAI
# client = OpenAI(api_key="", base_url="https://api.deepseek.com")
# api_message = [{"role": "system", "content": "你是Windows11操作引导者"}]

# 载入OmniParser相关权重
yolo_model = get_yolo_model(model_path='../OmniParser/weights/icon_detect/model.pt')  # 根据需要修改模型权重位置
caption_model_processor = get_caption_model_processor(
    model_name="florence2", 
    model_name_or_path="../OmniParser/weights/icon_caption_florence"  # 根据需要修改模型权重位置
)


def OmnRec(
    image_source: np.ndarray,
    box_threshold: float = 0.05,
    iou_threshold: float = 0.1,
    use_paddleocr: bool = True,
    imgsz: int = 640
):
    """
    部分参考自 https://zhuanlan.zhihu.com/p/24590537352
    OmniParser解析图像image_source
    """
    image = Image.fromarray(image_source)

    # 配置绘制边界框的参数
    box_overlay_ratio = image.size[0] / 3200
    draw_bbox_config = {
        'text_scale': 0.8 * box_overlay_ratio,
        'text_thickness': max(int(2 * box_overlay_ratio), 1),
        'text_padding': max(int(3 * box_overlay_ratio), 1),
        'thickness': max(int(3 * box_overlay_ratio), 1),
    }

    # OCR处理
    ocr_bbox_rslt, is_goal_filtered = check_ocr_box(
        image,
        display_img=False,
        output_bb_format='xyxy',
        goal_filtering=None,
        text_threshold=0.6,  # 0.9
    )
    text, ocr_bbox = ocr_bbox_rslt

    # 获取标记后的图片和解析内容
    dino_labled_img, parsed_content_list = get_som_labeled_img(
        image,
        yolo_model,
        BOX_TRESHOLD=box_threshold,
        output_coord_in_ratio=True,
        ocr_bbox=ocr_bbox,
        draw_bbox_config=draw_bbox_config,
        caption_model_processor=caption_model_processor,
        ocr_text=text,
        iou_threshold=iou_threshold,
        imgsz=imgsz,
    )

    h, w = image_source.shape[:2]
    for i, v in enumerate(parsed_content_list):
        bbox = parsed_content_list[i]['bbox']
        parsed_content_list[i]['bbox'] = [int(w*bbox[0]), int(h*bbox[1]), int(w*(bbox[2]-bbox[0])), int(h*(bbox[3]-bbox[1]))]

    return {
        "labeled_image": dino_labled_img,
        "parsed_content_list": parsed_content_list,
    }

def cv_show(img=None, img_name='default'):
    """
    显示numpy图像，主要用来debug
    """
    cv2.imshow(img_name, img)
    cv2.waitKey(0)
    cv2.destroyAllWindows()

def calculate_ssim_color(imageA, imageB):
    """
    计算ssim，计算两张图像的相似度
    """
    scores = []
    for i in range(3):  # 对每个通道（B, G, R）分别计算
        score, _ = ssim(imageA[:, :, i], imageB[:, :, i], full=True)
        scores.append(score)
    
    avg_score = np.mean(scores)
    similarity_percentage = (avg_score + 1) * 50
    
    return similarity_percentage

def findNClick(temp_img, clicks, button, bias=(0, 0), interval=0.5, time_lim=999999999,
                tpmatch_threshold=0.8, area=(0, 0, 1920, 1080),
                rgb_threshold=-1, after_sleep=0, verb_code=-1):
    """
    模板匹配+点击操作，点击模版图像出现的位置

    找模版图像temp_img，找到就点鼠标的button，点clicks次，点击位置是temp_img位置的中心点+偏移bias，找图时最多等待time_lim秒，超时就退出，点击后等待after_sleep秒
    模板匹配的识别区域是area，采用灰度图匹配，相似度阈值为tpmatch_threshold，彩色图可选择多一个颜色相似度阈值rgb_threshold
    """
    pos = OPT.is_exist(temp_img=temp_img, interval=interval, time_lim=time_lim,
                        tpmatch_threshold=tpmatch_threshold, area=area,
                        is_rgb=rgb_threshold>0, rgb_threshold=rgb_threshold)
    if pos is None:
        if verb_code > 0:
            print(verb_code, "failure")
        return None
    (x1, y1), (x2, y2) = pos
    if clicks > 0:
        pyautogui.click(x=area[0]+(x1+x2)//2+bias[0], y=area[1]+(y1+y2)//2+bias[1], clicks=clicks, button=button)
    if after_sleep > 0:
        time.sleep(after_sleep)
    if verb_code > 0:
        print(verb_code, "success")
    return pos

def findNMove(temp_img, bias=(0, 0), move_time=0.5, interval=0.5, time_lim=999999999,
                tpmatch_threshold=0.8, area=(0, 0, 1920, 1080),
                rgb_threshold=-1, after_sleep=0, verb_code=-1):
    """
    模板匹配+鼠标移动，移动鼠标到模版图像出现的位置

    找模版图像temp_img，找到就移动鼠标，移动的位置是temp_img位置的中心点+偏移bias，找图时最多等待time_lim秒，超时就退出，点击后等待after_sleep秒
    模板匹配的识别区域是area，采用灰度图匹配，相似度阈值为tpmatch_threshold，彩色图可选择多一个颜色相似度阈值rgb_threshold
    """
    pos = OPT.is_exist(temp_img=temp_img, interval=interval, time_lim=time_lim,
                        tpmatch_threshold=tpmatch_threshold, area=area,
                        is_rgb=rgb_threshold>0, rgb_threshold=rgb_threshold)
    if pos is None:
        if verb_code > 0:
            print(verb_code, "failure")
        return None
    (x1, y1), (x2, y2) = pos
    pyautogui.moveTo(x=area[0]+(x1+x2)//2+bias[0], y=area[1]+(y1+y2)//2+bias[1], duration=move_time)
    if after_sleep > 0:
        time.sleep(after_sleep)
    if verb_code > 0:
        print(verb_code, "success")
    return pos

def openConsoler():
    """
    打开控制台，显示输出
    """
    pos = OPT.is_exist(temp_img=consolerAddIcon, interval=0.5, time_lim=2,
                        tpmatch_threshold=0.9, area=(0, 0, SCREEN_WIDTH, SCREEN_HEIGHT))
    if pos is not None:
        return True
    pos2 = OPT.is_exist(temp_img=consolerIcon, interval=0.5, time_lim=2,
                        tpmatch_threshold=0.9, area=(0, SCREEN_HEIGHT-50, SCREEN_WIDTH, 50))
    if pos2 is None:
        return False
    x, y = (pos2[0][0]+pos2[1][0])//2, SCREEN_HEIGHT-50+(pos2[0][1]+pos2[1][1])//2
    pyautogui.click(x=x, y=y, clicks=1, button='left')
    return True

def closeConsoler():
    """
    最小化控制台
    """
    pos = OPT.is_exist(temp_img=consolerAddIcon, interval=0.5, time_lim=2,
                        tpmatch_threshold=0.9, area=(0, 0, SCREEN_WIDTH, SCREEN_HEIGHT))
    if pos is None:
        return False
    pos2 = OPT.is_exist(temp_img=consolerWinControlIcon, interval=0.5, time_lim=2,
                        tpmatch_threshold=0.9, area=(pos[0][0], pos[0][1]-5, SCREEN_WIDTH-pos[0][0], pos[1][1]-pos[0][1]+10))
    if pos2 is None:
        return False
    x, y = pos[0][0]+pos2[0][0]+16, pos[0][1]+pos2[0][1]+7
    pyautogui.click(x=x, y=y, clicks=1, button='left')
    return True

def openWenXiaoBai():
    """
    打开浏览器使用deepseek（这里用问小白，也可以换成其他平台的，但是要更换对应的模版图像，而且注意要单开浏览器给deepseek）
    """
    if findNClick(xiaobaiAiIcon, 0, 'left', time_lim=2, area=(0, 0, SCREEN_WIDTH, 50), tpmatch_threshold=0.8, after_sleep=0.5, verb_code=0):
        return True
    findNClick(edgeIcon, 1, 'left', time_lim=2, area=(0, SCREEN_HEIGHT-50, SCREEN_WIDTH, 50), tpmatch_threshold=0.8, verb_code=0)
    if findNClick(xiaobaiAiIcon, 1, 'left', time_lim=2, area=(0, 0, SCREEN_WIDTH, 50), tpmatch_threshold=0.8, after_sleep=0.5, verb_code=0):
        return True
    findNClick(wenXiaoBaiEdgeIcon, 1, 'left', time_lim=2, area=(2*SCREEN_WIDTH//5, SCREEN_HEIGHT-220, 7*SCREEN_WIDTH//20, 80), tpmatch_threshold=0.8, after_sleep=0.5, verb_code=0)
    return findNClick(xiaobaiAiIcon, 1, 'left', time_lim=2, area=(0, 0, SCREEN_WIDTH, 50), tpmatch_threshold=0.8, after_sleep=0.5, verb_code=0) is not None

def minimizeXiaoBai():
    """
    最小化deepseek浏览器窗口（这里用问小白，也可以换成其他平台的，单开deepseek的浏览器）
    """
    if findNClick(xiaobaiAiIcon, 0, 'left', time_lim=2, area=(0, 0, SCREEN_WIDTH, 50), tpmatch_threshold=0.95, rgb_threshold=0.95, after_sleep=0.5, verb_code=0) is None:
        return True
    return findNClick(broseWinControllIcon, 1, 'left', (-56, 0), time_lim=2, area=(0, 0, SCREEN_WIDTH, 50), tpmatch_threshold=0.95, rgb_threshold=0.95, after_sleep=0.5, verb_code=0) is not None

def enableXiaoBaiNet():
    """
    使用联网功能
    """
    focus_area = (SCREEN_WIDTH//4, SCREEN_HEIGHT//2-80, 5*SCREEN_WIDTH//8, SCREEN_HEIGHT//2+30)
    if findNClick(xiaobaiNetEnableIcon, 0, 'left', time_lim=2, area=focus_area, tpmatch_threshold=0.95, rgb_threshold=0.95, after_sleep=0.5, verb_code=0) is not None:
        return True
    return findNClick(xiaobaiNetDisableIcon, 1, 'left', time_lim=2, area=focus_area, tpmatch_threshold=0.95, rgb_threshold=0.95, after_sleep=0.5, verb_code=0) is not None

def disableXiaoBaiNet():
    """
    关闭联网功能（使其允许上传图片）
    """
    focus_area = (SCREEN_WIDTH//4, SCREEN_HEIGHT//2-80, 5*SCREEN_WIDTH//8, SCREEN_HEIGHT//2+30)
    if findNClick(xiaobaiNetDisableIcon, 0, 'left', time_lim=2, area=focus_area, tpmatch_threshold=0.9, verb_code=0) is not None:
        return True
    return findNClick(xiaobaiNetEnableIcon, 1, 'left', time_lim=2, area=focus_area, tpmatch_threshold=0.9, verb_code=0) is not None

# def getNextOrders(input_text: str):
#     """
#     调用api使用deepseek获取指令
#     """
#     response = None
#     api_message.append({"role": "user", "content": inputText})
#     try:
#         response = client.chat.completions.create(
#             model="deepseek-chat",  # chat reasoner
#             messages=api_message,
#             stream=False
#         )
#     except Exception as e:
#         if hasattr(e, 'response'):
#             print(f"Status Code: {e.response.status_code}")
#             print(f"Raw Response: {e.response.text}")  # 打印原始错误信息
#         else:
#             print(f"Error: {e}")
#         return
#     answer = response.choices[0].api_message.content
#     api_message.append({"role": "assistant", "content": answer})
#     orders = answer.split('== 执行 ==')[-1].split('\n')[1:]
#     orders = [order.split(';')[0] for order in orders]
#     return orders, answer

def newChannel2():
    """
    开一个新的对话
    """
    while True:
        open_res = openWenXiaoBai()
        print("openWenXiaoBai:", ('success' if open_res else 'failure'))
        if open_res:
            break
        time.sleep(1)

    findNClick(newXiaoBaiIcon, 1, 'left', time_lim=2, area=(0, 0, 400, 500), tpmatch_threshold=0.8, after_sleep=0.5, verb_code=0)
    return findNClick(xiaoBaiNewPageIcon, 0, 'left', time_lim=2, area=(SCREEN_WIDTH//4, 0, 3*SCREEN_WIDTH//4, SCREEN_HEIGHT//2), tpmatch_threshold=0.8, after_sleep=0.5, verb_code=0) is not None

def getNextOrders2(input_text: str, show_img=False):
    """
    单开浏览器白嫖deepseek获取指令
    """
    while True:
        open_res = openWenXiaoBai()
        print("openWenXiaoBai:", ('success' if open_res else 'failure'))
        if open_res:
            break
        time.sleep(1)

    focus_area = (SCREEN_WIDTH//4, SCREEN_HEIGHT//2-80, 5*SCREEN_WIDTH//8, SCREEN_HEIGHT//2+30)
    findNClick(sendIcon, 1, 'left', (-150, -80), time_lim=2, area=focus_area, tpmatch_threshold=0.8, after_sleep=0.5, verb_code=3)
    pyperclip.copy(input_text)
    pyautogui.hotkey('ctrl', 'v')
    if show_img:
        while True:
            open_res = disableXiaoBaiNet()
            print("disableXiaoBaiNet:", ('success' if open_res else 'failure'))
            if open_res:
                break
            time.sleep(1)
        findNClick(uploadImageIcon, 1, 'left', time_lim=2, area=focus_area, tpmatch_threshold=0.8, after_sleep=0.5, verb_code=4)
        findNClick(curStateImage, 2, 'left', time_lim=2, area=(0, 0, 1200, 715), tpmatch_threshold=0.8, after_sleep=0.5, verb_code=5)
    else:
        while True:
            open_res = enableXiaoBaiNet()
            print("enableXiaoBaiNet:", ('success' if open_res else 'failure'))
            if open_res:
                break
            time.sleep(1)
    findNClick(send2Icon, 1, 'left', time_lim=60, area=focus_area, tpmatch_threshold=0.85, after_sleep=0.5, verb_code=6)

    if show_img:
        findNClick(crossLastImgIcon, 1, 'left', time_lim=3, area=focus_area, tpmatch_threshold=0.9, after_sleep=0.5, verb_code=10)
    findNClick(msgBoxIcon, 1, 'left', time_lim=3, area=focus_area, tpmatch_threshold=0.8, after_sleep=0.5, verb_code=7)
    findNMove(sendIcon, (-120, -260), move_time=0.2, time_lim=999, area=focus_area, tpmatch_threshold=0.8, after_sleep=0.5, verb_code=8)
    pyautogui.scroll(-800)
    time.sleep(0.2)
    while findNClick(copyIcon, 1, 'left', time_lim=1.5, area=focus_area, tpmatch_threshold=0.8, after_sleep=0.2, verb_code=9) is None:
        pyautogui.scroll(-800)
        time.sleep(0.2)

    print("minimizeXiaoBai:", ('success' if minimizeXiaoBai() else 'failure'))

    answer = pyperclip.paste()
    orders = answer.split('== 执行 ==')[-1].split('\r\n')[1:]  # 指令流开始
    orders = [order.split(';')[0] for order in orders]
    return orders, answer


def run():
    task_text = "打开bilibili，给up主`极客湾Geekerwan`的最近一条动态点赞"  # 任务描述
    end_text = "点赞任务完成"  # 任务结束语

    task_text = f"{task_text}，完成所有任务后你输出`{end_text}`来让机器人停止"  # 总体任务描述

    # 指令格式说明
    order_format_text = "指令发送需要遵从以下格式：" +\
        "\n\n开始指令：指令发送开始是单独一行发送文本`== 执行 ==`" +\
        "\n\n鼠标操作指令：鼠标操作名$图标元素代码(横向偏移,纵向偏移)（偏移坐标可选）$图标元素代码（可选）;" +\
        "\n其中鼠标操作名有：鼠标移动、左键双击、左键单击、左键按住、左键松开、鼠标左键滑动、右键点击、右键按住、右键松开" +\
        "\n比如：`鼠标移动$1;`代表鼠标移动到序号1元素的位置，`左键单击$1;`代表鼠标左键单击序号1元素的位置，`左键单击$1(-80,0);`代表鼠标左键单击离序号1元素的位置偏移(-80,0)像素的位置，`左键双击$1;`代表鼠标左键双击序号1元素的位置，`鼠标左键滑动$1$2;`代表鼠标点击序号1元素的位置滑动到序号2元素的位置再松开" +\
        "\n\n键盘操作指令：键盘操作名$xxx（某段文本或单个按键字符串，**按键字符串使用pyautogui的KEYBOARD_KEYS**）$xxx（可选，只用于组合键操作给定额外的按键字符串，如有需要可以继续加`$组合键`，不限数量）;" +\
        "\n其中键盘操作名有：输入、键入、组合键入" +\
        "\n比如：`输入$你好;`代表使用键盘输入字符串`你好`，`键入$enter;`代表键盘按下回车键，`组合键入$ctrl$a$d;`代表组合键（热键）依次按下ctrl键和a键和d键" +\
        "\n\n等待指令：等待$x秒;" +\
        "\n比如：`等待$0.5秒;`代表等待0.5秒再执行下一个指令（如果有的话）" +\
        "\n\n图像指令：请求发送图像;" +\
        "\n\n解释：该指令用于机器人给你发送当前画面反馈，你需要一次机器人就只会返回一次，该指令效果不会延续到下一轮对话中"

    # 注意事项说明
    notation_text = "注意事项和推荐操作：" +\
        "\n1. **开始指令往后，仅允许发送鼠标指令、键盘指令、等待指令和图像指令，禁止任何非指令内容**，格式必须严格匹配，单条指令以分号`;`结束：" +\
        "\n   - 单行（单条）操作指令不能加任何前缀、后缀、补充性说明、自然语言陈述性描述和非提到的格式内容" +\
        "\n2. **在桌面时，禁止从底部任务栏打开浏览器**：" +\
        "\n   - 如果判断当前画面在Windows桌面，并且想点击浏览器时，禁止点击底部任务栏上的浏览器去打开，你只能选择点击桌面上的图标打开，因此你要考虑坐标位置做出选择" +\
        "\n3. **页面状态验证原则**：" +\
        "\n   - 跳转中断机制：当执行含键入~enter;、左键单击链接类元素等可能引发页面跳转的指令后，必须强制终止指令链，即使后续有逻辑上连续的操作。（注意，这很重要！！！）" +\
        "\n4. **关于操作**：" +\
        "\n   - 你只需想普通人一样思考怎么操作我给出的元素，不需要给出过于复杂的操作指令" +\
        "\n   - 禁止导致多次页面跳转的连续操作指令，这将导致操作失控，即相邻的指令执行在不同的画面上（但你不知道下一页面的内容），这是危险操作" +\
        "\n   - 着重考虑鼠标操作是否正确，比如单击是否需要换成双击（尤其是点击桌面上的图标时）" +\
        "\n   - 当画面重复次数过多时，反思是不是之前的操作间隔太短，尝试慢一点操作（间隔稍微放大一点）" +\
        "\n5. **元素内容描述反馈**：" +\
        "\n   - 元素内容描述都是基于OCR扫描出来，不一定正确，比如字母`Q`可能是`放大镜`图像元素（常见于搜索图标）" +\
        "\n   - 当你需要结合图像和给出的元素一起分析时，你需要利用图像指令" +\
        "\n   - 当文本描述中找不到你想要交互的元素时，你可以结合给出的屏幕图像和已有的各元素位置大小，在图像上找到你想操作的元素，结合最近的已有元素在此基础上用坐标偏移表达出该元素的位置" +\
        "\n6. **推荐与界面互动尝试以理解当前环境**：" +\
        "\n   - 给出的元素及其描述可能并不准确，如果你确定应该操作该元素附近的位置，可以尝试加上位置偏移" +\
        "\n   - 鼠标停留在某个元素上方查看元素是否有反应（即只使用`鼠标移动-XXX`和`等待-X秒`指令，当操作次数过多时推荐该做法查看该元素是否是正确对象）" +\
        "\n   - 上一操作执行不成功时，查找其它可能的类似操作或想其它解决方式（即避免与上一操作重复，避免出现循环现象）" +\
        "\n   - 善于利用每个元素的坐标，以此了解每个元素的大小，以及当前画面的布局" +\
        "\n   - 善于思考当前的处境和反思以往的操作" +\
        "\n7. **关于异常**：" +\
        "\n   - 网络连接不可能有异常，网络有异常，机器人不可能继续对话，你可以假设一切功能都是正常的，只是你的给的操作指令有问题，或者你的操作指令没有成功执行"+\
        "\n   - 任何软件都不可能有异常，流程之外的异常都不需要你排除和解决，当前画面不如期只能是单纯的操作指令不到位的问题"+\
        "\n   - 键盘操作指令是最难出错的，一般出错都是鼠标操作问题"

    # 首次对话前缀语
    prefix_text = "我是机器人，负责操作Windows11上的鼠标和键盘，我不能理解自然语言和解释性操作，只能接收规定格式的指令，同时我也不能保证每次操作都一定会准确完成（因为我提供的画面元素分析结果不一定准确），" + \
        "\n你是Windows11操作引导者，你需要一步一步给出当前需要做什么（若干条指令，每条指令给出等待时间），其中需要注意的是，你不能一次性发送很多指令执行完当前任务，并且你必须保证发送操纵鼠标相关的指令时，操作的画面元素在当前提供的画面内，" +\
        "\n每一步完成后我都会返回当前屏幕的关键操作元素的描述，描述以`序号`+`描述`的形式按行逐个列出来，同时还会给出当前画面和上一画面的相似度（以供参考是否正确完成，**大于95%时可能需要思考有别与上一步的操作跳出循环**），" +\
        "\n然后你需要结合当前信息分析当前页面，给出分析的过程，然后判断接下来需要做什么，再发送指令（在开始指令发送后，只能发送指令）。" +\
        f"\n\n{order_format_text}" +\
        f"\n\n{notation_text}" +\
        "\n\n每次开始指令之前先思考当前页面的状态和之前的操作，同时仔细思考给出的注意事项和推荐操作。" +\
        f"\n\n现在给出任务：{task_text}" +\
        "\n\n当前画面图标元素的描述如下，其中坐标格式为(元素左上角横坐标, 元素左上角纵坐标, 元素宽度，元素高度)，单位是像素且以屏幕左上角为原点，请遵守注意事项和推荐操作给出下一步指令："

    isFirst = True
    isOrderError = False
    errorOrderString = None
    consolerShowMsg = True
    show_img = False
    closeConsoler()

    imgIndex = 0
    sameRepeatTime = 0
    round_times = 0
    while True:
        round_times += 1
        if round_times % 16 == 0:  # 尝试超过16轮就换新
            while True:
                new_channel_res = newChannel2()
                print("newChannel2:", ('success' if new_channel_res else 'failure'))
                if new_channel_res:
                    break
                time.sleep(1)
            time.sleep(1)
            print("minimizeXiaoBai:", ('success' if minimizeXiaoBai() else 'failure'))
            isFirst = True
            isOrderError = False
            errorOrderString = None
            sameRepeatTime = 0
            round_times = 0
            time.sleep(1)
                
        if isOrderError:
            input_text = f"**指令格式错误：最近一次的操作指令序列中包含错误格式的指令，机器人未能识别**:\n{errorOrderString}\n\n回顾指令格式规范：\n{order_format_text}"
        else:
            # 识别画面元素
            img = fast_capture_full(0, 0, SCREEN_WIDTH, SCREEN_HEIGHT)
            Image.fromarray(img).save(f"logs/A{imgIndex}.png")
            Image.fromarray(img).save(f"imageSendCache/0.png")
            res = OmnRec(img)
            Image.fromarray(res['labeled_image']).save(f"logs/B{imgIndex}.png")

            # 完善prompt 尽可能地提供有用信息
            input_text = prefix_text if isFirst else "继续分析并提供指令，同时遵守注意事项和推荐操作，当前画面图标元素的描述如下，其中坐标格式为(元素左上角横坐标, 元素左上角纵坐标, 元素宽度，元素高度)，单位是像素且以屏幕左上角为原点："
            isFirst = False
            for i, v in enumerate(res['parsed_content_list']):
                input_text += f"\n{i}.  类型:`{v['type']}` 描述:`{v['content']}` 坐标:`{v['bbox']}`"
            if imgIndex > 0 and round_times > 1:
                sim = calculate_ssim_color(img, np.array(Image.open(f"logs/A{imgIndex-1}.png")))
                input_text += f"\n\n> 当前画面与上一画面相似度：{sim}%"
                sameRepeatTime += 1
                if sim > 95:
                    if sameRepeatTime > 3:
                        input_text += f"\n> 画面连续相似次数为{sameRepeatTime}次！"
                    else:
                        input_text += "\n> 你下一步可能需要做出不一样的操作"
                else:
                    sameRepeatTime = 0
                if round_times > 3 and round_times % 3 == 0:
                    input_text += f"\n\n请注意思考：当前已经是第{round_times}次操作了，当前流程是否按预期进行？是否是哪里出了问题呢？是否陷入循环呢？如果有循环怎样跳出呢？有没有结合元素坐标分析当前画面的布局？指令之间的间隔时间会不会太短而没执行成功？回顾注意事项：\n{notation_text}\n\n回顾任务：\n{task_text}\n\n然后继续给出指令"
            imgIndex += 1

        # 调用LLM 获取执行指令 控制台显示接下来的操作
        print('\n', input_text)
        orders, answer = getNextOrders2(input_text, show_img=show_img)
        if consolerShowMsg:
            openConsoler()
        print('\n', answer, '\n\n', orders)
        if consolerShowMsg:
            time.sleep(5)
            closeConsoler()
            time.sleep(1)

        # 检查指令前缀
        isOrderError = False
        errorOrderString = None
        opt_pos_list = []
        log_ipl = []
        try:
            for order in orders:
                errorOrderString = f"{order};"
                order = order.strip().split('$')
                if order[0] not in ("鼠标移动", "左键单击", "左键双击", "输入", "键入", "组合键入", "等待", "请求发送图像"):
                    break
                if order[0] in ("鼠标移动", "左键单击", "左键双击"):
                    eid = order[1]
                    rect_bias = (0, 0)
                    if '(' in eid:
                        eid = eid.split('(')
                        rect_bias = eid[-1][:-1].split(',')
                        rect_bias = (int(rect_bias[0].strip()), int(rect_bias[1].strip()))
                        eid = eid[0].strip()
                    eid = int(eid)
                    rect = res['parsed_content_list'][eid]['bbox']
                    opt_pos_list.append((rect[0] + rect[2] // 2 + rect_bias[0], rect[1] + rect[3] // 2 + rect_bias[1]))
                    log_ipl.append((order, opt_pos_list[-1]))
                elif order[0] in ('等待',):
                    wait_time = float(order[1][:-1])
                elif order[0] in ('图像指令',):
                    isOrderError = True
                    break
        except Exception as e:
            isOrderError = True
            print(f"发生了一个未预料的异常: {e}")
        if isOrderError:
            continue
        
        print("=" * 10)
        print(log_ipl)
        print("=" * 10)

        # 解析指令并执行，待完善
        show_img = False
        for order in orders:
            order = order.strip().split('$')
            if order[0] == end_text:
                return
            if order[0] == "鼠标移动":
                opt_pos = opt_pos_list.pop(0)
                print("鼠标移动: ", order, opt_pos)
                pyautogui.moveTo(x=opt_pos[0], y=opt_pos[1], duration=0.3)
            elif order[0] == "左键单击":
                opt_pos = opt_pos_list.pop(0)
                print("左键单击: ", order, opt_pos)
                pyautogui.click(x=opt_pos[0], y=opt_pos[1])
            elif order[0] == "左键双击":
                opt_pos = opt_pos_list.pop(0)
                print("左键双击: ", order, opt_pos)
                pyautogui.click(x=opt_pos[0], y=opt_pos[1], clicks=2)
            elif order[0] == "输入":
                pyperclip.copy(order[1])
                pyautogui.hotkey('ctrl', 'v')
                print("输入: ", order)
            elif order[0] == "键入":
                pyautogui.press(order[1])
                print("键入: ", order)
            elif order[0] == "组合键入":  # 待完善，切屏的快捷键还不能实现，可以尝试分解成 按下-按下-松开-松开 的形式
                pyautogui.hotkey(*order[1:])
                print("组合键入: ", order)
            elif order[0] == "等待":
                time.sleep(float(order[1][:-1]))
                print("等待: ", order)
            elif order[0] == "请求发送图像":
                show_img = True
                print("请求发送图像: ", order)
            else:
                break
            time.sleep(0.5)


if __name__ == '__main__':
    run()
