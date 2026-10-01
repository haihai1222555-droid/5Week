@echo off`ndocker run --rm -it --cap-add=SYS_PTRACE --security-opt seccomp=unconfined -v "%cd%:/work" -w /work memdbg
