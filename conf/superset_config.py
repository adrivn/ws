# Superset specific config
ROW_LIMIT = 5000

# Flask App Builder configuration
# Your App secret key will be used for securely signing the session cookie
# and encrypting sensitive information on the database
# Make sure you are changing this key for your deployment with a strong key.
# Alternatively you can set it with `SUPERSET_SECRET_KEY` environment variable.
# You MUST set this for production environments or the server will not refuse
# to start and you will see an error in the logs accordingly.
SECRET_KEY = "E50)]V.P20QwXOq5@m=e:pQ/qYDpzuaQG0t(?w?U+..57,*Rm}w3r+7qCoOJyl[x"

# The SQLAlchemy connection string to your database backend
# This connection defines the path to the database that stores your
# superset metadata (slices, connections, tables, dashboards, ...).
# Note that the connection information to connect to the datasources
# you want to explore are managed directly in the web UI
SQLALCHEMY_DATABASE_URI = (
    "sqlite:///N:/CoralHudson/1. AM/8. Wholesale Channel/.code/supersetdb.db"
)

# Flask-WTF flag for CSRF
WTF_CSRF_ENABLED = True
# Add endpoints that need to be exempt from CSRF protection
WTF_CSRF_EXEMPT_LIST = []
# A CSRF token that expires in 1 year
WTF_CSRF_TIME_LIMIT = 60 * 60 * 24 * 365

# Set this API key to enable Mapbox visualizations
MAPBOX_API_KEY = "pk.eyJ1IjoiYW1vbnRlc2ciLCJhIjoiY2xuYTNvbWJhMDAwZTJrbzNpaDdqNjE3cCJ9.qQH-CYzrVYNKrLyoIjSBuQ"

# Export
WEBDRIVER_BASEURL = "http://127.0.0.1:8088/"
WEBDRIVER_TYPE = "chrome"
SCREENSHOT_LOCATE_WAIT = 40
SCREENSHOT_LOAD_WAIT = 60
