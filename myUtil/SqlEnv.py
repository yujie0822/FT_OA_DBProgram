import os
import sys
stdout = sys.stdout
stdin = sys.stdin
stderr = sys.stderr
reload(sys)
sys.setdefaultencoding('utf-8')
sys.stdout = stdout
sys.stdin = stdin
sys.stderr = stderr
os.environ['NLS_LANG'] = 'SIMPLIFIED CHINESE_CHINA.UTF8'

# default values
DEFAULT_OA_USER = "oadb"
DEFAULT_OA_PASSWORD = "oracle"
DEFAULT_JDE_USER = "jdetest"
DEFAULT_JDE_PASSWORD = "jdetest"
DEFAULT_OA_CONNECT_STRING = "192.168.0.89:1521/OADB"
DEFAULT_JDE_CONNECT_STRING = "192.168.0.238:1521/E1DB"

# values that will be used are the default values unless environment variables
# have been set as noted above
OA_USER = DEFAULT_OA_USER
OA_PASSWORD = DEFAULT_OA_PASSWORD
JDE_USER = DEFAULT_JDE_USER
JDE_PASSWORD = DEFAULT_JDE_PASSWORD
OA_CONNECT_STRING = DEFAULT_OA_CONNECT_STRING
JDE_CONNECT_STRING = DEFAULT_JDE_CONNECT_STRING
# calculated values based on the values above
MAIN_OA_CONNECT_STRING = "%s/%s@%s" % (OA_USER, OA_PASSWORD, OA_CONNECT_STRING)
MAIN_JDE_CONNECT_STRING = "%s/%s@%s" % (JDE_USER, JDE_PASSWORD, JDE_CONNECT_STRING)
