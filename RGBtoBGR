import matplotlib.pyplot as plt
from PIL import Image

from sam3 import build_sam3_image_model
from sam3.model.box_ops import box_xywh_to_cxcywh
from sam3.model.sam3_image_processor import Sam3Processor
from sam3.visualization_utils import draw_box_on_image, normalize_bbox, plot_results
import cv2
import io
import numpy as np


checkpoint_path = "f:/sam3/sam3.pt"
bpe_path = "f:/sam3/assets/bpe_simple_vocab_16e6.txt.gz"
model = build_sam3_image_model(bpe_path=bpe_path, checkpoint_path=checkpoint_path)

image_path = "f:/sam3/assets/images/test_image.jpg"
image = Image.open(image_path)
# width, height = image.size
processor = Sam3Processor(model, confidence_threshold=0.5)
inference_state = processor.set_image(image)

processor.reset_all_prompts(inference_state)
inference_state = processor.set_text_prompt(state=inference_state, prompt="shoe")

img0 = Image.open(image_path)
plot_results(img0, inference_state)

plt.axis("off")
buf = io.BytesIO()
plt.savefig(buf, format='png', bbox_inches='tight', pad_inches=0)
buf.seek(0)

# 转换为NumPy数组
img_pil = Image.open(buf)
img_array = np.array(img_pil)
# PIL格式 (RGB) 转 OpenCV格式 (BGR)
img_cv2 = cv2.cvtColor(img_array, cv2.COLOR_RGB2BGR)
cv2.imshow("result_img", img_cv2)
cv2.waitKey(0)
# 关闭buffer
plt.close()

# plt.show()
