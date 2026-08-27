from dify_plugin import Plugin, DifyPluginEnv
import os

os.environ["LOAD_FROM_DIFY_PLUGIN"] = "1"

plugin = Plugin(DifyPluginEnv(MAX_REQUEST_TIMEOUT=120))

if __name__ == '__main__':
    plugin.run()
