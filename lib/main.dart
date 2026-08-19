import 'dart:convert';
import 'dart:io';
import 'package:excel/excel.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/material.dart';
import 'package:intl/intl.dart';
import 'package:provider/provider.dart';
import 'package:shared_preferences/shared_preferences.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  final store=Store(); await store.load();
  runApp(ChangeNotifierProvider.value(value:store,child:const App()));
}
class Client{
  String id,name,artist,phone,email; double share,advance;
  Client({required this.id,required this.name,this.artist='',this.phone='',this.email='',this.share=.7,this.advance=0});
  Map<String,dynamic> j()=>{'id':id,'name':name,'artist':artist,'phone':phone,'email':email,'share':share,'advance':advance};
  factory Client.f(Map x)=>Client(id:x['id'],name:x['name'],artist:x['artist']??'',phone:x['phone']??'',email:x['email']??'',share:(x['share'] ?? 0.7).toDouble(),advance:(x['advance']??0).toDouble());
}
class Song{
  String isrc,name,artist,clientId;
  Song({required this.isrc,required this.name,this.artist='',this.clientId=''});
  Map<String,dynamic> j()=>{'isrc':isrc,'name':name,'artist':artist,'clientId':clientId};
  factory Song.f(Map x)=>Song(isrc:x['isrc'],name:x['name'],artist:x['artist']??'',clientId:x['clientId']??'');
}
class Royalty{
  String dsp,month,song,isrc; double income,admin,royalty;
  Royalty({this.dsp='',this.month='',this.song='',this.isrc='',this.income=0,this.admin=0,this.royalty=0});
  Map<String,dynamic> j()=>{'dsp':dsp,'month':month,'song':song,'isrc':isrc,'income':income,'admin':admin,'royalty':royalty};
  factory Royalty.f(Map x)=>Royalty(dsp:x['dsp']??'',month:x['month']??'',song:x['song']??'',isrc:x['isrc']??'',income:(x['income']??0).toDouble(),admin:(x['admin']??0).toDouble(),royalty:(x['royalty']??0).toDouble());
}
class Payment{
  String clientId,date,ref,note; double amount;
  Payment({required this.clientId,this.date='',this.ref='',this.note='',this.amount=0});
  Map<String,dynamic> j()=>{'clientId':clientId,'date':date,'ref':ref,'note':note,'amount':amount};
  factory Payment.f(Map x)=>Payment(clientId:x['clientId'],date:x['date']??'',ref:x['ref']??'',note:x['note']??'',amount:(x['amount']??0).toDouble());
}
class AdvanceRecovery{
  String clientId,date,note; double amount;
  AdvanceRecovery({required this.clientId,this.date='',this.note='',this.amount=0});
  Map<String,dynamic> j()=>{'clientId':clientId,'date':date,'note':note,'amount':amount};
  factory AdvanceRecovery.f(Map x)=>AdvanceRecovery(clientId:x['clientId'],date:x['date']??'',note:x['note']??'',amount:(x['amount']??0).toDouble());
}
class Store extends ChangeNotifier{
  List<Client> clients=[]; List<Song> songs=[]; List<Royalty> rows=[]; List<Payment> payments=[]; List<AdvanceRecovery> recoveries=[];
  final money=NumberFormat.currency(locale:'en_IN',symbol:'₹');
  Future<void> load()async{final p=await SharedPreferences.getInstance();
    clients=(jsonDecode(p.getString('clients')??'[]') as List).map((x)=>Client.f(x)).toList();
    songs=(jsonDecode(p.getString('songs')??'[]') as List).map((x)=>Song.f(x)).toList();
    rows=(jsonDecode(p.getString('rows')??'[]') as List).map((x)=>Royalty.f(x)).toList();
    payments=(jsonDecode(p.getString('payments')??'[]') as List).map((x)=>Payment.f(x)).toList();
    recoveries=(jsonDecode(p.getString('recoveries')??'[]') as List).map((x)=>AdvanceRecovery.f(x)).toList(); notifyListeners();}
  Future<void> save()async{final p=await SharedPreferences.getInstance();
    await p.setString('clients',jsonEncode(clients.map((x)=>x.j()).toList()));
    await p.setString('songs',jsonEncode(songs.map((x)=>x.j()).toList()));
    await p.setString('rows',jsonEncode(rows.map((x)=>x.j()).toList()));
    await p.setString('payments',jsonEncode(payments.map((x)=>x.j()).toList()));
    await p.setString('recoveries',jsonEncode(recoveries.map((x)=>x.j()).toList())); notifyListeners();}
  Client? clientFor(String isrc){final ss=songs.where((x)=>x.isrc.trim().toUpperCase()==isrc.trim().toUpperCase());if(ss.isEmpty)return null;final cc=clients.where((x)=>x.id==ss.first.clientId);return cc.isEmpty?null:cc.first;}
  double clientGross(String id)=>rows.where((r)=>clientFor(r.isrc)?.id==id).fold(0.0,(a,r)=>a+r.royalty*(clientFor(r.isrc)?.share??0));
  double recovered(String id)=>recoveries.where((x)=>x.clientId==id).fold(0.0,(a,x)=>a+x.amount);
  double paid(String id)=>payments.where((x)=>x.clientId==id).fold(0.0,(a,x)=>a+x.amount);
  double advanceBalance(String id){final c=clients.firstWhere((x)=>x.id==id);return (c.advance-recovered(id)).clamp(0,double.infinity);}
  double payable(String id)=>(clientGross(id)-recovered(id)).clamp(0,double.infinity);
  double outstanding(String id)=>(payable(id)-paid(id)).clamp(0,double.infinity);
  double get totalRoyalty=>rows.fold(0.0,(a,r)=>a+r.royalty);
  double get totalClient=>clients.fold(0.0,(a,c)=>a+clientGross(c.id));
  double get totalLabel=>totalRoyalty-totalClient;
  double get totalOutstanding=>clients.fold(0.0,(a,c)=>a+outstanding(c.id));
  Future<void> importPDL()async{
    final f=await FilePicker.platform.pickFiles(type:FileType.custom,allowedExtensions:['xlsx']); if(f==null||f.files.single.path==null)return;
    final book=Excel.decodeBytes(await File(f.files.single.path!).readAsBytes()); final sh=book.tables.values.first; if(sh.rows.isEmpty)return;
    final h=sh.rows.first.map((c)=>(c?.value??'').toString().trim().toLowerCase()).toList();
    int i(String n)=>h.indexOf(n.toLowerCase()); String v(List<Data?> r,String n){final z=i(n);return z>=0&&z<r.length?(r[z]?.value??'').toString():'';}
    double d(List<Data?> r,String n)=>double.tryParse(v(r,n).replaceAll(',','').replaceAll('₹',''))??0;
    rows.clear();
    for(final r in sh.rows.skip(1)){final isrc=v(r,'ISRC');if(isrc.isNotEmpty)rows.add(Royalty(dsp:v(r,'DSP Name'),month:v(r,'Month of Royalty'),song:v(r,'Song Name'),isrc:isrc,income:d(r,'Income'),admin:d(r,'Admin_Exp'),royalty:d(r,'Royalty Paid to PDL member')));}
    await save();
  }
}
class App extends StatelessWidget{const App({super.key});@override Widget build(BuildContext c)=>MaterialApp(debugShowCheckedModeBanner:false,title:'Audioholic Label Manager',theme:ThemeData(colorSchemeSeed:Colors.indigo,useMaterial3:true),home:const Home());}
class Home extends StatefulWidget{const Home({super.key});@override State<Home> createState()=>_HomeState();}
class _HomeState extends State<Home>{int i=0;final pages=[const Dashboard(),const ClientsPage(),const SongsPage(),const ImportPage(),const AccountsPage(),const PaymentsPage()];final titles=['Dashboard','Clients','Songs & ISRC','Import PDL Excel','Royalty Accounts','Payments'];@override Widget build(BuildContext c)=>Scaffold(appBar:AppBar(title:Text(titles[i])),body:pages[i],bottomNavigationBar:NavigationBar(selectedIndex:i,onDestinationSelected:(x)=>setState(()=>i=x),destinations:const[NavigationDestination(icon:Icon(Icons.dashboard),label:'Home'),NavigationDestination(icon:Icon(Icons.people),label:'Clients'),NavigationDestination(icon:Icon(Icons.music_note),label:'Songs'),NavigationDestination(icon:Icon(Icons.upload_file),label:'Import'),NavigationDestination(icon:Icon(Icons.account_balance_wallet),label:'Accounts'),NavigationDestination(icon:Icon(Icons.payments),label:'Payments')]));}
class Dashboard extends StatelessWidget{const Dashboard({super.key});@override Widget build(BuildContext c){final s=c.watch<Store>();Widget card(String t,double v,IconData i)=>Card(child:Padding(padding:const EdgeInsets.all(16),child:Column(crossAxisAlignment:CrossAxisAlignment.start,children:[Icon(i),const Spacer(),Text(t),Text(s.money.format(v),style:const TextStyle(fontSize:20,fontWeight:FontWeight.bold))])));return Padding(padding:const EdgeInsets.all(10),child:GridView.count(crossAxisCount:2,childAspectRatio:1.15,children:[card('Total Royalty',s.totalRoyalty,Icons.currency_rupee),card('Client Earnings',s.totalClient,Icons.person),card('Label Earnings',s.totalLabel,Icons.business),card('Outstanding',s.totalOutstanding,Icons.pending_actions),card('Imported Rows',s.rows.length.toDouble(),Icons.table_rows),card('Clients',s.clients.length.toDouble(),Icons.groups)]));}}
class ClientsPage extends StatelessWidget{const ClientsPage({super.key});Future<void> add(BuildContext c)async{final a=TextEditingController(),b=TextEditingController(),ar=TextEditingController(),sh=TextEditingController(text:'70'),ad=TextEditingController(text:'0');final ok=await showDialog<bool>(context:c,builder:(x)=>AlertDialog(title:const Text('Add Client'),content:SingleChildScrollView(child:Column(children:[TextField(controller:a,decoration:const InputDecoration(labelText:'Client ID')),TextField(controller:b,decoration:const InputDecoration(labelText:'Client Name')),TextField(controller:ar,decoration:const InputDecoration(labelText:'Artist')),TextField(controller:sh,keyboardType:TextInputType.number,decoration:const InputDecoration(labelText:'Revenue Share %')),TextField(controller:ad,keyboardType:TextInputType.number,decoration:const InputDecoration(labelText:'Recoverable Advance'))])),actions:[TextButton(onPressed:()=>Navigator.pop(x),child:const Text('Cancel')),FilledButton(onPressed:()=>Navigator.pop(x,true),child:const Text('Save'))]));if(ok==true){final s=c.read<Store>();s.clients.add(Client(id:a.text.trim(),name:b.text.trim(),artist:ar.text.trim(),share:(double.tryParse(sh.text)??70)/100,advance:double.tryParse(ad.text)??0));await s.save();}}@override Widget build(BuildContext c){final s=c.watch<Store>();return Scaffold(floatingActionButton:FloatingActionButton(onPressed:()=>add(c),child:const Icon(Icons.add)),body:ListView.builder(itemCount:s.clients.length,itemBuilder:(c,i){final x=s.clients[i];return ListTile(title:Text(x.name),subtitle:Text('${x.id} • ${(x.share*100).toStringAsFixed(0)}% • Outstanding ${s.money.format(s.outstanding(x.id))}'),trailing:Text(s.money.format(s.payable(x.id)));}));}}
class SongsPage extends StatelessWidget {
  const SongsPage({super.key});

  Future<void> add(BuildContext c) async {
    final isrcController = TextEditingController();
    final songController = TextEditingController();
    String? clientId;

    final ok = await showDialog<bool>(
      context: c,
      builder: (context) {
        final store = c.read<Store>();
        return AlertDialog(
          title: const Text('Add Song / ISRC'),
          content: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              TextField(
                controller: isrcController,
                decoration: const InputDecoration(labelText: 'ISRC'),
              ),
              TextField(
                controller: songController,
                decoration: const InputDecoration(labelText: 'Song Name'),
              ),
              DropdownButtonFormField<String>(
                items: store.clients
                    .map(
                      (client) => DropdownMenuItem(
                        value: client.id,
                        child: Text('${client.id} - ${client.name}'),
                      ),
                    )
                    .toList(),
                onChanged: (value) => clientId = value,
                decoration: const InputDecoration(labelText: 'Client'),
              ),
            ],
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.pop(context),
              child: const Text('Cancel'),
            ),
            FilledButton(
              onPressed: () => Navigator.pop(context, true),
              child: const Text('Save'),
            ),
          ],
        );
      },
    );

    if (ok == true) {
      final store = c.read<Store>();
      store.songs.add(
        Song(
          isrc: isrcController.text.trim(),
          name: songController.text.trim(),
          clientId: clientId ?? '',
        ),
      );
      await store.save();
    }
  }

  @override
  Widget build(BuildContext c) {
    final store = c.watch<Store>();
    return Scaffold(
      floatingActionButton: FloatingActionButton(
        onPressed: () => add(c),
        child: const Icon(Icons.add),
      ),
      body: ListView(
        children: store.songs
            .map(
              (song) => ListTile(
                title: Text(song.name),
                subtitle: Text(
                  'ISRC: ${song.isrc} • Client: ${song.clientId}',
                ),
              ),
            )
            .toList(),
      ),
    );
  }
}
class ImportPage extends StatelessWidget{const ImportPage({super.key});@override Widget build(BuildContext c)=>Padding(padding:const EdgeInsets.all(24),child:Column(crossAxisAlignment:CrossAxisAlignment.start,children:[const Text('Import PDL Distributor Excel',style:TextStyle(fontSize:22,fontWeight:FontWeight.bold)),const SizedBox(height:12),const Text('Expected columns: DSP Name, Month of Royalty, Song Name, ISRC, Income, Admin_Exp, Royalty Paid to PDL member.'),const SizedBox(height:24),FilledButton.icon(onPressed:()async{try{await c.read<Store>().importPDL();if(c.mounted)ScaffoldMessenger.of(c).showSnackBar(const SnackBar(content:Text('PDL report imported successfully')));}catch(e){if(c.mounted)ScaffoldMessenger.of(c).showSnackBar(SnackBar(content:Text('Import failed: $e')));}},icon:const Icon(Icons.upload_file),label:const Text('Select Excel File')),const SizedBox(height:16),Consumer<Store>(builder:(c,s,_)=>
Text('Imported rows: ${s.rows.length}'))]));}
class AccountsPage extends StatelessWidget{const AccountsPage({super.key});@override Widget build(BuildContext c){final s=c.watch<Store>();return ListView.builder(itemCount:s.clients.length,itemBuilder:(c,i){final x=s.clients[i];return Card(child:ListTile(title:Text(x.name),subtitle:Text('Gross: ${s.money.format(s.clientGross(x.id))} • Advance Balance: ${s.money.format(s.advanceBalance(x.id))}\\nPaid: ${s.money.format(s.paid(x.id))}'),trailing:Column(mainAxisAlignment:MainAxisAlignment.center,children:[const Text('OUTSTANDING'),Text(s.money.format(s.outstanding(x.id)),style:const TextStyle(fontWeight:FontWeight.bold))]));});}}
class PaymentsPage extends StatelessWidget {
  const PaymentsPage({super.key});

  Future<void> add(BuildContext c, bool recovery) async {
    final amountController = TextEditingController();
    final referenceController = TextEditingController();
    final dateController = TextEditingController(
      text: DateFormat('yyyy-MM-dd').format(DateTime.now()),
    );
    String? clientId;

    final ok = await showDialog<bool>(
      context: c,
      builder: (context) {
        final store = c.read<Store>();
        return AlertDialog(
          title: Text(recovery ? 'Advance Recovery' : 'Client Payment'),
          content: Column(
            mainAxisSize: MainAxisSize.min,
            children: [
              DropdownButtonFormField<String>(
                items: store.clients
                    .map(
                      (client) => DropdownMenuItem(
                        value: client.id,
                        child: Text(client.name),
                      ),
                    )
                    .toList(),
                onChanged: (value) => clientId = value,
                decoration: const InputDecoration(labelText: 'Client'),
              ),
              TextField(
                controller: amountController,
                keyboardType: TextInputType.number,
                decoration: const InputDecoration(labelText: 'Amount'),
              ),
              TextField(
                controller: dateController,
                decoration: const InputDecoration(
                  labelText: 'Date YYYY-MM-DD',
                ),
              ),
              TextField(
                controller: referenceController,
                decoration: InputDecoration(
                  labelText: recovery ? 'Note' : 'UTR / Reference',
                ),
              ),
            ],
          ),
          actions: [
            TextButton(
              onPressed: () => Navigator.pop(context),
              child: const Text('Cancel'),
            ),
            FilledButton(
              onPressed: () => Navigator.pop(context, true),
              child: const Text('Save'),
            ),
          ],
        );
      },
    );

    if (ok == true && clientId != null) {
      final store = c.read<Store>();
      final amount = double.tryParse(amountController.text) ?? 0;

      if (recovery) {
        store.recoveries.add(
          AdvanceRecovery(
            clientId: clientId!,
            amount: amount,
            date: dateController.text,
            note: referenceController.text,
          ),
        );
      } else {
        store.payments.add(
          Payment(
            clientId: clientId!,
            amount: amount,
            date: dateController.text,
            ref: referenceController.text,
          ),
        );
      }
      await store.save();
    }
  }

  @override
  Widget build(BuildContext c) {
    final store = c.watch<Store>();

    return Scaffold(
      floatingActionButton: Column(
        mainAxisAlignment: MainAxisAlignment.end,
        children: [
          FloatingActionButton.extended(
            heroTag: 'recovery',
            onPressed: () => add(c, true),
            label: const Text('Recover Advance'),
          ),
          const SizedBox(height: 10),
          FloatingActionButton.extended(
            heroTag: 'payment',
            onPressed: () => add(c, false),
            label: const Text('Add Payment'),
          ),
        ],
      ),
      body: ListView(
        children: [
          const ListTile(
            title: Text(
              'Client Payment History',
              style: TextStyle(fontSize: 20, fontWeight: FontWeight.bold),
            ),
          ),
          ...store.payments.reversed.map(
            (payment) {
              final matches = store.clients
                  .where((client) => client.id == payment.clientId);
              final clientName = matches.isEmpty
                  ? payment.clientId
                  : matches.first.name;

              return ListTile(
                title: Text(clientName),
                subtitle: Text('${payment.date} • ${payment.ref}'),
                trailing: Text(store.money.format(payment.amount)),
              );
            },
          ),
        ],
      ),
    );
  }
}
