# THIS FILE IS GENERATED FROM PADDLEPADDLE SETUP.PY
#
full_version    = '2.1.2'
major           = '2'
minor           = '1'
patch           = '2'
rc              = '0'
istaged         = True
commit          = 'e04b66f2d272d68f77dcd94cb2956938475411d8'
with_mkl        = 'ON'

def show():
    if istaged:
        print('full_version:', full_version)
        print('major:', major)
        print('minor:', minor)
        print('patch:', patch)
        print('rc:', rc)
    else:
        print('commit:', commit)

def mkl():
    return with_mkl
