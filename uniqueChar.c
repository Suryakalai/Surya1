#include<stdio.h>
#include<string.h>
int main()
{
	char a[]="malayalam";
	int i,j;
	for(i=0;i<strlen(a);i++){
		for(j=i+1;j<strlen(a);j++){
			if(a[i]==a[j]){
				a[j]='0';
			}	
		}
		if(a[i]!='0'){
			printf("%c",a[i]);
		}
	}
}
