import urllib.request
import os

base_url = "https://releases.aspose.com/java/repo/com/aspose/aspose-words/25.12/"
filenames = ["aspose-words-25.12-jdk17.jar", "aspose-words-25.12.jar"]
output_dir = "lib"

for filename in filenames:
    url = base_url + filename
    output_path = os.path.join(output_dir, filename)
    print(f"Trying to download {url}...")
    try:
        urllib.request.urlretrieve(url, output_path)
        print(f"Successfully downloaded {filename}")
        break
    except Exception as e:
        print(f"Failed to download {filename}: {e}")
