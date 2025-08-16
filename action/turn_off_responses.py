import requests

def main():
      url = "https://script.google.com/macros/s/AKfycbzTn8ctnhs2ihvt9x4lZH5b-zGEugJ-k9H1i3iUBw-t2Avs56zULZGKZHRZrlYrHDXK/exec"
      params = {
            "formId": "13sjagi0j_7DeRgdLrD_cO-ZDigWYF2CqWCS-UNxfQBg"
      }

      response = requests.get(url, params=params)
      print(response.text)
     
if __name__ == "__main__":
      main()

