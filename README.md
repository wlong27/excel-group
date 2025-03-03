# excel-group

```bash
python -m venv venv
source venv/Scripts/activate
pip install streamlit pandas openpyxl
streamlit run main.py

docker build -t my-app-01 .
docker tag my-app-01:latest <username>/my-app-01:latest
docker push <username>/my-app-01:latest
```
