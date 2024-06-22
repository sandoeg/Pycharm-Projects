import requests
import json

# ARV BR sensor_id = 13955981
# Get Sensor Readings: https://tempstickapi.com/api/v1/sensor/sensor_id/readings?setting=today
# Get All Sensors: https://tempstickapi.com/api/v1/sensors/all
# Get Sensor: https://tempstickapi.com/api/v1/sensor/sensor_id
# Get All Alerts: https://tempstickapi.com/api/v1/alerts/all
# Get Current User: https://tempstickapi.com/api/v1/user
# Set Display Preferences: https://tempstickapi.com/api/v1/user/display-preferences?temp_pref=F&chart_fill=0

headers = {
    'X-API-KEY': 'aefb575f918c4e08493075cdfa62df3f5491d57e18ebb2ab2d'
}

# response = requests.get('https://tempstickapi.com/api/v1/user', headers=headers)

response = requests.post(
    "https://tempstickapi.com/api/v1/user/display-preferences?json={timezone:America/Chicago}&temp_pref=F", headers=headers)

x = json.loads(response.content)
y = json.dumps(x, indent=4)
print(y)
