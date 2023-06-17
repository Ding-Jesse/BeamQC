# import pickle
import pprint
import dill as pickle
import timeit
from item import beam
def save_pkl(data,tmp_file=''):
    with open(tmp_file,'wb') as tmp:
        pickle.dump(data,tmp)
def read_temp(tmp_file=''):
    with open(tmp_file,'rb') as tmp:
        data = pickle.load(tmp)
    # pprint.pprint(data)
    return data
def save_object_pkl(data_list,tmp_file=''):
    with open(tmp_file,'wb') as tmp:
        for data in data_list:
            pickle.dump(data,tmp)
def load_temp(filename):
    with open(filename, "rb") as f:
        while True:
            try:
                yield pickle.load(f)
            except EOFError:
                break
class Company():
    def __init__(self, name, value):
        self.name = name
        self.value = value
        self.person = ''
if __name__=='__main__':
    def f():
        save_pkl(l,r'test.pkl')
        print(read_temp(r'test.pkl'))
    l = [ beam.Beam(1,1,1)]*10
    print(timeit.timeit(f, number=10))
