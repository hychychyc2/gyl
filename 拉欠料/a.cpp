#include<cstdio>
#iclude<vector>
using namespace std;
vector <vector<int>> v={{'A','B','C'},{'a','b'},{'1','2'}};
void print(vector <vector<int>> & input){
    int  max_len=0;
    for(int i=0;i<input.size()){
        max_len=max(input[i].size(),max_len);
    }
    for(int i=0;i<max_len;i++){
        if(i&1){
            for(int j=input.size()-1;j>=0;j--){
                if(i<input[j].size())
                printf("%c",input[j][i]);
            }
        }else{
            for(int j=0;j<input.size();j++){
                if(i<input[j].size())
                printf("%c",input[j][i]);
            }
        }
    }
}
int main(){
  print(v);
}