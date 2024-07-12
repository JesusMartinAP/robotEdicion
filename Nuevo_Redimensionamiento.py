# %%
from tqdm import tqdm
from PIL import Image
import imutils
import cv2
import os 

# %%
import shutil
os.makedirs("Temp",exist_ok=True)
ruta=os.listdir("Input")
#ruta=[x for x in ruta if ("jpg" in str(x))]

# %%
def remove_transparency(path, bg_colour=(255, 255, 255)):
    try:
        im = Image.open(r"Input/"+path) 
        if im.mode in ('RGBA', 'LA') or (im.mode == 'P' and 'transparency' in im.info):
            alpha = im.convert('RGBA').split()[-1]
            bg = Image.new("RGBA", im.size, bg_colour + (255,))
            bg.paste(im, mask=alpha)
            bg.convert('RGB').save(r"Temp/"+path)
        else:
            im.convert('RGB').save(r"Temp/"+path)
    except Exception as e:
        print(f"Error {e} en {path}")

for x in tqdm(ruta):
    remove_transparency(x)

imagenes=os.listdir("Temp")

for imagen in tqdm(imagenes):
    # Load an image
    img = cv2.imread("Temp/"+imagen)
    # Get the dimensions of the image
    height, width, channels = img.shape
    if height>1400 or width>1400:
        medidas=[height,width]
        mayor=medidas.index(max(medidas))
        if mayor==0: ## la mayor medida es height
            img = imutils.resize(img,height=1400)
        if mayor==1:
            img = imutils.resize(img,width=1400)
        cv2.imwrite("Output/"+imagen,img)
    else:
        cv2.imwrite("Output/"+imagen,img)
shutil.rmtree("Temp")
