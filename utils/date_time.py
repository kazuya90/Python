from datetime import datetime

def get_current_datetime():
  now = datetime.now()
  return now.strftime("%Y%m%d%H%M")

if __name__ == "__main__":
  print(get_current_datetime())