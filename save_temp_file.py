import pickle
import pprint
def save_pkl(data,tmp_file=''):
    with open(tmp_file,'wb') as tmp:
        pickle.dump(data,tmp)
def read_temp(tmp_file=''):
    with open(tmp_file,'rb') as tmp:
        data = pickle.load(tmp)
    # pprint.pprint(data)
    return data

if __name__=='__main__':
    read_temp()