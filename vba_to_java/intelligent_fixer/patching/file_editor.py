"""Constrained JSON/text editor with optimistic concurrency checks."""
import json, hashlib

class FileEditor:
    @staticmethod
    def _hash(data):
        return hashlib.sha256(data).hexdigest()

    @staticmethod
    def _tokens(pointer):
        if not isinstance(pointer, str) or not pointer.startswith("/"):
            raise ValueError("JSON pointer must start with /")
        return [x.replace("~1", "/").replace("~0", "~") for x in pointer[1:].split("/")]

    def json_set(self, raw, pointer, old_value, new_value):
        obj=json.loads(raw.decode("utf-8-sig")); cur=obj
        tokens=self._tokens(pointer)
        for token in tokens[:-1]:
            cur=cur[int(token)] if isinstance(cur,list) else cur[token]
        last=tokens[-1]
        existing=cur[int(last)] if isinstance(cur,list) else cur[last]
        if existing != old_value:
            raise ValueError("Patch precondition failed: old_value mismatch")
        if isinstance(cur,list): cur[int(last)]=new_value
        else: cur[last]=new_value
        return (json.dumps(obj,indent=2,ensure_ascii=False)+"\n").encode("utf-8")

    def json_replace(self, raw, pointer, old_value, new_value):
        return self.json_set(raw,pointer,old_value,new_value)

    def json_add(self, raw, pointer, new_value):
        obj=json.loads(raw.decode("utf-8-sig")); cur=obj
        tokens=self._tokens(pointer)
        for token in tokens[:-1]:
            cur=cur[int(token)] if isinstance(cur,list) else cur[token]
        last=tokens[-1]
        if isinstance(cur,list):
            if last=="-": cur.append(new_value)
            else: cur.insert(int(last),new_value)
        else:
            if last in cur:
                raise ValueError("Patch precondition failed: target key already exists")
            cur[last]=new_value
        return (json.dumps(obj,indent=2,ensure_ascii=False)+"\n").encode("utf-8")

    def text_replace(self, raw, old_text, new_text):
        text=raw.decode("utf-8-sig")
        if not old_text or text.count(old_text)!=1:
            raise ValueError("old_text must match exactly once")
        encoded=text.replace(old_text,new_text,1).encode("utf-8")
        return (b"\xef\xbb\xbf"+encoded) if raw.startswith(b"\xef\xbb\xbf") else encoded

    def json_remove(self, raw, pointer, old_value):
        obj=json.loads(raw.decode("utf-8-sig")); cur=obj
        tokens=self._tokens(pointer)
        for token in tokens[:-1]:
            cur=cur[int(token)] if isinstance(cur,list) else cur[token]
        last=tokens[-1]
        existing=cur[int(last)] if isinstance(cur,list) else cur[last]
        if existing != old_value:
            raise ValueError("Patch precondition failed: old_value mismatch")
        if isinstance(cur,list): del cur[int(last)]
        else: del cur[last]
        return (json.dumps(obj,indent=2,ensure_ascii=False)+"\n").encode("utf-8")

    def file_replace(self, raw, new_content, suffix=""):
        if not isinstance(new_content, str):
            raise ValueError("new_content must be a string")
        if len(new_content.encode("utf-8")) > 100_000:
            raise ValueError("Whole-file replacement is limited to 100 KB; use a narrow patch")
        if suffix.lower() in {".json", ".ipynb"}:
            import json
            json.loads(new_content)
        encoded=new_content.encode("utf-8")
        return (b"\xef\xbb\xbf"+encoded) if raw.startswith(b"\xef\xbb\xbf") else encoded
