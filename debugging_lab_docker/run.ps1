docker run --rm -it --cap-add=SYS_PTRACE --security-opt seccomp=unconfined -v "${PWD}:/work" -w /work memdbg
