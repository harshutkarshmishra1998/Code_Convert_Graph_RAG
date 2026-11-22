"""Execute a reviewed patch through the transactional patch manager."""
from ..patching.patch_manager import PatchManager
class RepairExecutor:
    def __init__(self,project_root): self.manager=PatchManager(project_root)
    def apply(self,proposal):
        return self.manager.apply(proposal)
