import appModuleHandler
import globalPluginHandler


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		appModuleHandler.registerExecutableWithAppModule("spcsadm", "vismaAdmin")
		appModuleHandler.registerExecutableWithAppModule("spcsfkt", "vismaAdmin")
		appModuleHandler.registerExecutableWithAppModule("spcsfor", "vismaAdmin")

	def terminate(self, *args, **kwargs):
		super().terminate(*args, **kwargs)
		appModuleHandler.unregisterExecutable("spcsadm")
		appModuleHandler.unregisterExecutable("spcsfkt")
		appModuleHandler.unregisterExecutable("spcsfor")
