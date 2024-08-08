A9='June-24'
A8=' Achieved'
A7=' Target'
A6='Weekly Report Siri Suggi(1)'
A5='key.json'
A4='https://www.googleapis.com/auth/drive'
A3='https://spreadsheets.google.com/feeds'
A2='FixedColumn'
w='_'
v=enumerate
o=sum
m='15-21 July Margin\u200b'
j='MTD Margin\u200b'
i='w'
h='r'
e='Gross Margin'
d='15-21 Achieved\u200b'
c='15-21 Target\u200b'
b='-'
a=''
Z=zip
Y=range
W='YTD Margin\u200b'
V='Purchase Inventory (₹)'
T=str
S='Achievement YTD %\u200b'
R='Achieve YTD\u200b'
Q='Target YTD\u200b'
P='Achievement for Month %\u200b'
O='MTD Achieved June-24\u200b'
N='MTD Target July-24\u200b'
M='Achieved % \u200b'
L='Total Sales (₹)'
K=open
I=int
H='Sales Target (₹)'
G=len
D='---------------------------------'
C=round
B=print
A=float
import csv as J,glob as U,os as f,pandas as E,gspread as x
from oauth2client.service_account import ServiceAccountCredentials as y
import itertools as AA,requests as z,json
from datetime import datetime as F,timedelta as X
AB='TBD*.csv'
AC='SBD*.csv'
AD='CBD*.csv'
AE='CI_*.csv'
AF='CIB_E*.csv'
AG='SI_*.csv'
AH='SIB_E*.csv'
AI='TWC*.csv'
n=U.glob(AB)
p=U.glob(AC)
g=U.glob(AD)
q=U.glob(AE)
r=U.glob(AF)
s=U.glob(AG)
t=U.glob(AI)
u=U.glob(AH)
n.sort()
p.sort()
g.sort()
q.sort()
s.sort()
r.sort()
t.sort()
u.sort()
B(n)
B(p)
B(g)
B(q)
B(r)
B(s)
B(t)
B(u)
def A0(month):
	A=month
	if A in[1,3,5,7,8,10,12]:return 31
	elif A in[4,6,9,11]:return 30
	elif A==2:return 29
	else:raise ValueError('Invalid month number. Must be between 1 and 12.')
def AJ(file_path,columns_to_extract,new_headers,fixed_values):
	D=[];F=AA.cycle(fixed_values)
	with K(file_path,mode=h)as G:
		H=J.DictReader(G)
		for E in H:
			A={}
			for(C,I)in Z(columns_to_extract,new_headers):
				if C in E:A[I]=E[C]
				else:B(f"Column '{C}' not found in the CSV file.");return
			A[A2]=next(F);D.append(A)
	return D
def k(file_path):
	with K(file_path,mode=h)as A:B=J.DictReader(A);return[A for A in B]
def AW(file_path,columns):
	with K(file_path,mode=h)as A:B=J.DictReader(A);return[{A:B[A]for A in columns if A in B}for B in B]
def AK(data,output_file):
	with K(output_file,mode=i,newline=a)as B:A=J.DictWriter(B,fieldnames=data[0].keys());A.writeheader();A.writerows(data)
def l(file_path,data,headers):
	with K(file_path,mode=i)as A:B=J.DictWriter(A,fieldnames=headers);B.writerows(data)
def AX(file_path,data,headers):
	with K(file_path,mode=i,newline=a)as B:
		A=J.DictWriter(B,fieldnames=headers)
		for(C,D)in v(data):
			if C%13==0:A.writeheader()
			A.writerow(D)
def A1(file_path,header):
	B={}
	with K(file_path,mode=h)as E:
		F=J.DictReader(E)
		for D in F:
			C=D[header].strip();G=A(D[V].replace(',',a).strip())
			if C not in B:B[C]=0
			B[C]+=G
	return B
def AL(matching_files4):
	A1='lotDetails.rate';v='servicecharge';u='sellingprice';t='data';s='application/json';r='Content-Type';q='http://13.126.125.132:4000';h='gst';a='sale_value_xgst';Y='servicecharge_unit_xgst';W='sellingprice_unit_xgst';V='%Y-%m-%dT23:59:59Z';P='%Y-%m-%d';N='soldqty';M='product';H='invoicedate';G='%Y-%m-%dT00:00:00Z';A2=q;A9='{"query":"query suggi_GetSaleWiseProfit { suggi_getSaleWiseProfit { _id customerDetails { name onboarding_date  village { name }  pincode { code } cust_type GSTIN iscouponapplicable address  { city street  pincode } phone customer_uid} storeDetails { name address {   pincode } vertical territory { name zone { name } } } userDetails { name address {   pincode } } invoiceno invoicedate product { soldqty servicecharge extradiscount sellingprice gst purchaseProductDetails {   name   category {     name   }   manufacturer {     name   }   sub_category {     name   } } lotDetails {  discount rate id sellingprice invoice { supplier_ref invoiceno createddate invoicedate } cnStockproduct subqty transportCharges otherCharges landingrate } supplierDetails { name } storeTargetPerProduct } payment { card cash upi } grosstotal storeCost } } ","variables":{}}';AA={r:s};matching_files4.sort();i=n[1];c=f.path.basename(i);Q=c.split(w);J=Q[2];Q=J.split(b);O=Q[1];J=Q[0];d=T(2024);S='2024-'+O+b+J;J=T(I(J)+6);j=A0(I(O))
	if I(J)>j:
		J=T(I(J)-j);O=T(I(O)+1)
		if I(O)==13:O=T(I(O)-12);d=T(d+1)
	U=d+b+O+b+J;AB=z.post(A2,headers=AA,data=A9);AC=AB.json()[t]['suggi_getSaleWiseProfit'];e=E.DataFrame(AC)
	def AD(df,start_date_str,end_date_str):
		Q='gst%_unit';O='servicecharge_unit';L='sellingprice_unit';D=df;I=F.strptime(start_date_str,P);R=I.month;S=I.year;T=I;U=F(2024,R,1);Z=F(S,4,1);b=T.strftime(G);c=U.strftime(G);d=Z.strftime(G);J=(F.strptime(end_date_str,P)+X(days=1)-X(seconds=1)).strftime(V);e=[b,c,d];K=[]
		for f in e:D[H]=E.to_datetime(D[H]);I=E.to_datetime(f);J=E.to_datetime(J);g=D[(D[H]>=I)&(D[H]<=J)];B=g.explode([M]);B[L]=B[M].apply(lambda x:x[u]);B[O]=B[M].apply(lambda x:x[v]);B[N]=B[M].apply(lambda x:x[N]);B[Q]=B[M].apply(lambda x:x[h]);B[W]=B[L]/(1+B[Q]/100);B[Y]=B[O]/1.18;B[a]=B[N]*(B[W]+B[Y]);sum=B[a].sum();K.append(sum)
		i=[C(A(B)/100000,2)for B in K];return i
	def AE(start_date_str,end_date_str):
		H='Date';K=q;L='{"query":"query Suggi_StoreTarget {\\n  suggi_StoreTarget {\\n    Store\\n    TM\\n    Category\\n    Month\\n    Year\\n    Target\\n    Date\\n    Daily_Target\\n  }\\n}","variables":{}}';M={r:s};N=z.post(K,headers=M,data=L);O=N.json()[t]['suggi_StoreTarget'];B=E.DataFrame(O);D=F.strptime(start_date_str,P);Q=D.month;R=D.year;S=D;T=F(2024,Q,1);U=F(R,4,1);W=S.strftime(G);Y=T.strftime(G);Z=U.strftime(G);I=(F.strptime(end_date_str,P)+X(days=1)-X(seconds=1)).strftime(V);a=[W,Y,Z];J=[]
		for b in a:B[H]=E.to_datetime(B[H]);D=E.to_datetime(b);I=E.to_datetime(I);c=B[(B[H]>=D)&(B[H]<=I)];sum=c['Daily_Target'].sum();J.append(sum)
		d=[C(A(B)/100000,2)for B in J];return d
	def AF(df,start_date_str,end_date_str):
		S='actual_margin';R='purchase_value_xgst';Q='purchasingprice_unit';L='lotDetails.subqty';K='lotDetails.discount';D=df;D[H]=E.to_datetime(D[H]);I=F.strptime(start_date_str,P);U=I.month;Z=I.year;b=I;c=F(2024,U,1);d=F(Z,4,1);e=b.strftime(G);f=c.strftime(G);g=d.strftime(G);J=(F.strptime(end_date_str,P)+X(days=1)-X(seconds=1)).strftime(V);i=[e,f,g];O=[]
		for j in i:I=E.to_datetime(j);J=E.to_datetime(J);k=D[(D[H]>=I)&(D[H]<=J)];l=k.explode(M);B=E.json_normalize(l[M]);B[W]=B[u]/(1+B[h]/100);B[Y]=B[v]/(1+B[h]/100);B[a]=B[N]*(B[W]+B[Y]);B[Q]=B[A1];B[K]=B[K];B[R]=B[N]*(B[Q]-B[K]);B[S]=B[a]-(B[R]-B[N]*(B['lotDetails.cnStockproduct']/B[L])+B[N]*(B['lotDetails.transportCharges']/B[L])+B[N]*(B['lotDetails.otherCharges']/B[L]));O.append(B[S].sum())
		m=[C(A(B)/100000,2)for B in O];n=[T(A)+'%'for A in m];return n
	def AG(df,start_date_str,end_date_str):
		L='achived_purchase';B=df;B[H]=E.to_datetime(B[H]);D=F.strptime(start_date_str,P);O=D.month;Q=D.year;R=D;S=F(2024,O,1);T=F(Q,4,1);U=R.strftime(G);W=S.strftime(G);Y=T.strftime(G);J=(F.strptime(end_date_str,P)+X(days=1)-X(seconds=1)).strftime(V);Z=[U,W,Y];K=[]
		for a in Z:D=E.to_datetime(a);J=E.to_datetime(J);b=B[(B[H]>=D)&(B[H]<=J)];c=b.explode(M);I=E.json_normalize(c[M]);I[L]=I[N]*I[A1];K.append(I[L].sum())
		d=[C(A(B)/100000,2)for B in K];return d
	k=AD(e,S,U);l=AE(S,U);m=AF(e,S,U);o=AG(e,S,U);B(k);B(l);B(m);B(o);AH=[A3,A4];AI=y.from_json_keyfile_name(A5,AH);AJ=x.authorize(AI);AK=A6;AL=AJ.open(AK);R=AL.get_worksheet(16);K=['C2','F2','I2']
	for(D,L)in Z(K,k):R.update(values=[[L]],range_name=D)
	K=['B2','E2','H2']
	for(D,L)in Z(K,l):R.update(values=[[L]],range_name=D)
	K=['C5','F5','I5']
	for(D,L)in Z(K,m):R.update(values=[[L]],range_name=D)
	K=['C3','F3','I3']
	for(D,L)in Z(K,o):R.update(values=[[L]],range_name=D)
	c=f.path.basename(i);Q=c.split(w);p=Q[2].split(b);D=p[0]+b+T(I(p[0])+6);g=[];g.append(D+A7);g.append(D+A8);K=['B1','C1']
	for(D,L)in Z(K,g):R.update(values=[[L]],range_name=D)
def AM(matching_files2):
	V=matching_files2;V.sort();E=[]
	for X in V:B(X);E.append(k(X))
	I=[];D={c:0,d:0,M:0,m:0,N:0,O:0,P:0,j:0,Q:0,R:0,S:0,W:0}
	for F in Y(G(E[0])):
		J=C(A(E[1][F][H])/100000,2);Z=C(A(E[1][F][L])/100000,2);h=C(Z/J*100 if J!=0 else 0,2);i=C(A(E[1][F][e])/100000,2);K=C(A(E[0][F][H])/100000,2);a=C(A(E[0][F][L])/100000,2);n=C(a/K*100 if K!=0 else 0,2);o=C(A(E[0][F][e])/100000,2);T=C(A(E[2][F][H])/100000,2);b=C(A(E[2][F][L])/100000,2);p=C(b/T*100 if T!=0 else 0,2);q=C(A(E[2][F][e])/100000,2);f={c:J,d:Z,M:h,m:i,N:K,O:a,P:n,j:o,Q:T,R:b,S:p,W:q};I.append(f)
		for g in D:D[g]+=f[g]
	D[M]=A(D[d])/A(D[c]);D[P]=A(D[O])/A(D[N]);D[S]=A(D[R])/A(D[Q]);r={A:f"{B:.2f}"for(A,B)in D.items()};I.append(r);U='15-21';u=A9;s='Suggi_2 Report.csv';t=[U+' Target\u200b',U+' Achieved\u200b',M,U+' July Margin\u200b',N,O,P,j,Q,R,S,W];l(s,I,t)
def AN(matching_files3):
	V=matching_files3;U='Total Sales\u200b';T='Sales Target\u200b';V.sort();D=[]
	for X in V:B(X);D.append(k(X))
	F=[]
	for E in Y(G(D[0])):
		I=C(A(D[1][E][H])/100000,2);Z=C(A(D[1][E][L])/100000,2);c=C(Z/I*100 if I!=0 else 0,2);J=C(A(D[0][E][H])/100000,2);a=C(A(D[0][E][L])/100000,2);d=C(a/J*100 if J!=0 else 0,2);K=C(A(D[2][E][H])/100000,2);b=C(A(D[2][E][L])/100000,2);f=C(b/K*100 if K!=0 else 0,2);g=C(A(D[2][E][e])/100000,2);F.append({T:I,U:Z,M:c,N:J,O:a,P:d,Q:K,R:b,S:f,W:g})
		if G(F)==13 or G(F)==27:F.append({T:0,U:0,M:0,N:0,O:0,P:0,Q:0,R:0,S:0,W:0})
	h=[T,U,M,N,O,P,Q,R,S,W];i='Suggi_3 Report.csv';l(i,F,h)
def AO(matching_files4):
	J=matching_files4;J.sort();D=[]
	for K in J:B(K);D.append(k(K))
	I=[];F={c:0,d:0,M:0,m:0,N:0,O:0,P:0,j:0,Q:0,R:0,S:0,W:0}
	for E in Y(G(D[0])):
		T={c:C(A(D[0][E][H])/100000,2),d:C(A(D[0][E][L])/100000,2),M:C(A(D[0][E][L])/A(D[0][E][H])*100 if A(D[0][E][H])!=0 else 0,2),m:C(A(D[0][E][e])/100000,2),N:C(A(D[2][E][H])/100000,2),O:C(A(D[2][E][L])/100000,2),P:C(A(D[2][E][L])/A(D[2][E][H])*100 if A(D[2][E][H])!=0 else 0,2),j:C(A(D[2][E][e])/100000,2),Q:C(A(D[1][E][H])/100000,2),R:C(A(D[1][E][L])/100000,2),S:C(A(D[1][E][L])/A(D[1][E][H])*100 if A(D[1][E][H])!=0 else 0,2),W:C(A(D[1][E][e])/100000,2)};I.append(T)
		for U in F:F[U]+=T[U]
	F[M]=A(F[d])/A(F[c]);F[P]=A(F[O])/A(F[N]);F[S]=A(F[R])/A(F[Q]);V={A:f"{B:.2f}"for(A,B)in F.items()};I.append(V);a='15-21';b=A9;X='Suggi_4 Report.csv';Z=[c,d,M,m,N,O,P,j,Q,R,S,W];l(X,I,Z)
def AP(matching_files5):
	E=matching_files5;E.sort()
	for H in E:B(H)
	M='\ufeffCategory';I=[A1(A,M)for A in E];F=set()
	for N in I:F.update(N.keys())
	D=[];F=sorted(F)
	for O in F:A=[C(A.get(O,0)/100000,2)for A in I];D.append(A)
	P=[o(B[A]for B in D)for A in Y(0,G(E))];Q=P;D.append(Q)
	for A in D:R=o(A[0:]);A.append(R);S=A[3];L=A[4];T=C(S/L,2)if L!=0 else 0;A.append(T)
	U='Suggi_5 Report.csv'
	with K(U,mode=i,newline=a)as H:V=J.writer(H);V.writerows(D)
def AQ(matching_files6):
	H=matching_files6;H.sort();D=[]
	for I in H:B(I);D.append(k(I))
	E=[];F={V:0}
	for L in Y(G(D[0])):
		J={V:C(A(D[0][L][V])/100000,2)};E.append(J)
		for K in F:F[K]+=J[K]
	M={A:f"{B:.2f}"for(A,B)in F.items()};E.append(M);N='Suggi_6 Report.csv';O=[V];l(N,E,O)
def AR(matching_files7):
	E=matching_files7;E.sort()
	for H in E:B(H)
	M='\ufeffterritory';I=[A1(A,M)for A in E];F=set()
	for N in I:F.update(N.keys())
	D=[];F=sorted(F)
	for O in F:A=[C(A.get(O,0)/100000,2)for A in I];D.append(A)
	P=[o(B[A]for B in D)for A in Y(0,G(E))];Q=P;D.append(Q)
	for A in D:R=o(A[0:]);A.append(R);S=A[3];L=A[4];T=C(S/L,2)if L!=0 else 0;A.append(T)
	U='Suggi_7 Report.csv'
	with K(U,mode=i,newline=a)as H:V=J.writer(H);V.writerows(D)
def AS(matching_files8):
	R='Achieved %';Q='Target';P='extracted_data.csv';F=matching_files8;E='Customer';F.sort();B(F);S=F[0];T=['\ufeffcustomer','territory'];U=[E,'Territory'];V=['5819','4367','5852','4303'];H=f.path.join(S)
	if f.path.exists(H):
		I=AJ(H,T,U,V)
		if I:AK(I,P)
		else:B(f"File not found: {H}")
	C=[];D=P
	if f.path.exists(D):
		with K(D,mode=h,encoding='utf-8-sig')as W:
			X=J.DictReader(W)
			for Z in X:C.append(Z)
	L=[]
	for M in Y(G(C)):N=A(C[M][E]);O=A(C[M][A2]);b=N/O;L.append({Q:O,E:N,R:b})
	c='Suggi_8 Report.csv'
	with K(c,mode=i,newline=a)as D:d=[Q,E,R];e=J.DictWriter(D,fieldnames=d);e.writerows(L)
def AT(matching_files9):
	H=matching_files9;H.sort();D=[]
	for I in H:B(I);D.append(k(I))
	E=[];F={V:0}
	for L in Y(G(D[0])):
		J={V:C(A(D[0][L][V])/100000,2)};E.append(J)
		for K in F:F[K]+=J[K]
	M={A:f"{B:.2f}"for(A,B)in F.items()};E.append(M);N='Suggi_9 Report.csv';O=[V];l(N,E,O)
B(D)
AL(g)
B(D)
AM(n)
B(D)
AN(p)
B(D)
AO(g)
B(D)
AP(q)
B(D)
AQ(r)
B(D)
AR(s)
B(D)
AS(t)
B(D)
AT(u)
B(D)
B('Reports Generated')
B(D)
def AU():
	g.sort();D=g[1];E=f.path.basename(D);F=E.split(w);B=F[2].split(b);A=T(I(B[0])+6);C=A0(I(B[1]))
	if I(A)>C:A=T(I(A)-C)
	G=B[0]+' - '+A;return G
def AV():
	Y='Suggi_third_report';b=[A3,'https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive.file',A4];c=y.from_json_keyfile_name(A5,b);d=x.authorize(c);R=d.open(A6);e='*Report.csv';S=U.glob(e);T=['Suggi_second_report',Y,'Suggi_fourth_report','Suggi_fifth_report','Suggi_sixth_report','Suggi_seventh_report','Suggi_eigthth_report','Suggi_ninth_report'];S.sort();B(D);B(f"Updated Suggi_first_report values to Suggi_first_report");B(D)
	for(C,N)in v(S):
		with K(N,h)as f:
			g=J.reader(f);F=list(g);B(D);B(f"Updating {N} to {T[C]}");E=R.worksheet(T[C]);E.batch_clear(['B2:ZZ1000']);V=AU();O=[];O.append(V+A7);O.append(V+A8)
			if C<3:
				i=['B1','C1']
				for(M,C)in Z(i,O):E.update(values=[[C]],range_name=M)
			if F:
				m=G(F);H=G(F[0])if F else 0;L=[]
				for(W,j)in v(F):
					if C==1 and(W==12 or W==25):L.append([a]*H)
					L.append([A(B)for B in j])
				k=f"B2:{chr(65+H)}{G(L)+1}";E.update(values=L,range_name=k);B(f"Updated {G(L)} rows and {H} columns");B(D)
			else:B(f"{N} is a empty file")
	E=R.worksheet(Y);X=1;P=[];P.append(I(15));P.append(I(29));Q=E.row_values(X)
	for M in P:
		if Q:H=G(Q);l=f"A{M}:{chr(64+H)}{M}";E.update(l,[Q])
		else:B(f"Header in third report are not added {X}")
	B(D);B('Report is updated successfully');B(D)
AV()