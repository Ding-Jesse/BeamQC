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
    print(read_temp(r'D:\Desktop\BeamQC\TEST\INPUT\temp-0107-2023-01-07-11-09-temp.pkl'))