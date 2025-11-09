import appModuleHandler
import os
import globalPluginHandler
import api
from typing import Iterable, List, Any
import speech
from logHandler import log


# This code has been written because NVDA sometimes announces the text "ruta" (window) when moving up and down in one of the grid
# controls that has full row selection. It's very very vey very annoying and it slows down your navigation radically.
# I don't know why this happens but it's probably due to window messages/focus events firing. It can be as many as ten times announced so this
# was a welcome workaround. Yes, you can hear the difference because there's a little delay instead of all the "ruta" but that's ok in parity.
TARGET_EXE = "spcsadm.exe"
TEXT_TO_SUPPRESS = "ruta"


def _is_target_context() -> bool:
	f = api.getFocusObject()
	if not f:
		return False
	# App check
	am = getattr(f, "appModule", None)
	app_ok = False
	if am:
		appPath = (getattr(am, "appPath", "") or "")
		app_ok = os.path.basename(appPath).lower() == TARGET_EXE.lower()
		if not app_ok:
			appName = (getattr(am, "appName", "") or "").lower()
			app_ok = appName in {TARGET_EXE.lower(), TARGET_EXE.lower().removesuffix(".exe")}
	return app_ok


def _item_text(item: Any) -> str | None:
	"""Best-effort: speech sequences can contain strings and command objects."""
	if isinstance(item, str):
		return item
	# Many speech command objects expose .text (e.g. TextCommand).
	txt = getattr(item, "text", None)
	return txt if isinstance(txt, str) else None

def _filter_sequence(seq: Iterable[Any]) -> List[Any]:
	if not _is_target_context():
		return list(seq)
	out: List[Any] = []
	suppress = TEXT_TO_SUPPRESS.lower()
	for item in seq:
		txt = _item_text(item)
		if txt is not None and txt.strip().lower() == suppress:
			# Skip this piece entirely.
			continue


		out.append(item)
	return out


class GlobalPlugin(globalPluginHandler.GlobalPlugin):

	def __init__(self, *args, **kwargs):
		super().__init__(*args, **kwargs)
		# Defensive: confirm we're bound to NVDA's real module
		try:
			modpath = getattr(speech, "__file__", "")
			if not modpath or "nvda" not in modpath.lower():
				raise RuntimeError(f"Unexpected 'speech' module import: {modpath}")
		except Exception:
			# If this triggers, there’s still a name collision.
			pass
		self._reg = speech.filter_speechSequence.register(_filter_sequence)
		appModuleHandler.registerExecutableWithAppModule("spcsadm", "vismaAdmin")
		appModuleHandler.registerExecutableWithAppModule("spcsfkt", "vismaAdmin")
		appModuleHandler.registerExecutableWithAppModule("spcsfor", "vismaAdmin")


	def terminate(self, *args, **kwargs):
		super().terminate(*args, **kwargs)
		appModuleHandler.unregisterExecutable("spcsadm")
		appModuleHandler.unregisterExecutable("spcsfkt")
		appModuleHandler.unregisterExecutable("spcsfor")
		try:
			self._reg.remove()
		except:
			pass

