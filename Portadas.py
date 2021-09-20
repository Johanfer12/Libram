from bing_image_downloader import downloader
import os

path = "E://pruebando"
buscar = "Paul the Peddler; Or, The Fortunes of a Young Street Merchant by Alger, Horatio, Jr."

downloader.download(str(buscar) + " cover book", limit=1,  output_dir=path, adult_filter_off=True, force_replace=False, timeout=60)

busca = buscar.replace(':', ';')

os.remove(path + "//" + busca +".jpg")

os.rename(path + "//Image_1.jpg",path + "//" + busca +".jpg")
