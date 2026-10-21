"""Substitutos mínimos de System.IO (File, Directory) e System.Array para rodar
as funções do TXT de keynotes fora do Revit (CPython). FAIL liga falhas simuladas."""
import os, shutil, sys, types

FAIL = {}


class File(object):
    @staticmethod
    def Exists(p):
        return os.path.isfile(p)

    @staticmethod
    def Delete(p):
        if os.path.exists(p):
            os.remove(p)

    @staticmethod
    def ReadAllBytes(p):
        with open(p, "rb") as f:
            return bytearray(f.read())

    @staticmethod
    def WriteAllBytes(p, data):
        if FAIL.get("write"):
            raise IOError("disco cheio (simulado)")
        if any(c in os.path.basename(p) for c in '<>|"?*'):
            raise IOError("Illegal characters in path.")
        with open(p, "wb") as f:
            f.write(bytes(bytearray(data)))

    @staticmethod
    def Copy(a, b, overwrite=False):
        if os.path.exists(b) and not overwrite:
            raise IOError("exists")
        if FAIL.get("copy_backup") and "_obsoletos" in b:
            raise IOError("cópia recusada (simulado)")
        if FAIL.get("locked") and a.endswith(".tmp"):
            raise IOError("The process cannot access the file because it is being used by "
                          "another process (simulado)")
        shutil.copy(a, b)
        if FAIL.get("corrupt_backup") and "_obsoletos" in b:
            with open(b, "ab") as f:
                f.write(b"x")

    @staticmethod
    def Move(a, b):
        if os.path.exists(b):
            raise IOError("exists")
        os.rename(a, b)


class Directory(object):
    @staticmethod
    def Exists(p):
        return os.path.isdir(p)

    @staticmethod
    def CreateDirectory(p):
        if FAIL.get("mkdir"):
            raise IOError("acesso negado (simulado)")
        if not os.path.isdir(p):
            os.makedirs(p)


class _ArrT(object):
    def __getitem__(self, t):
        return lambda data: bytearray(data)


System = types.ModuleType("System")
System.IO = types.ModuleType("System.IO")
System.IO.File, System.IO.Directory = File, Directory
System.Array, System.Byte = _ArrT(), int
sys.modules["System"], sys.modules["System.IO"] = System, System.IO
