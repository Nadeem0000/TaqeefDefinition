import * as React from 'react';
import { useRef } from "react"; 
import styles from './TaqeefDefinitions.module.scss';
import { ITaqeefDefinitionsProps } from './ITaqeefDefinitionsProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Web } from "@pnp/sp/webs";
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from 'jquery';
import { ImageLoadState } from 'office-ui-fabric-react';
// import Select from 'react-select';
import { Autocomplete } from '@material-ui/lab';
import TextField from '@material-ui/core/TextField';
import ReactTooltip from "react-tooltip";
import { Items } from '@pnp/sp/items';
import * as moment from 'moment';
import { Markup } from 'interweave';



SPComponentLoader.loadCss('https://cdn.rawgit.com/brianreavis/selectize.js/master/dist/css/selectize.css');
SPComponentLoader.loadCss(`https://fonts.googleapis.com/css?family=Roboto:300,400,500,700`);           
SPComponentLoader.loadCss(`https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css`);   
SPComponentLoader.loadCss(`https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css`); 
SPComponentLoader.loadCss(`https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/CSS/DefinitionStyle.css?v=2.1`);
SPComponentLoader.loadCss(`https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/CSS/responsive.css?v=1.4`);

export interface ITaqeefDefinitionsState {  
  value:any[];  
  Data:any[];
  Product:any[];
  ProductFilter:any[];
  isClearable: boolean;
  isSearchable:boolean;
  Productvariant:any[];
  ProductCategory:any[];
  ProductSelected:string;
  DepartmentSelected:string;
  DivisionSelected:string;
  TagsSelected:string;
  CurrentlySelectedSearchTab:string;
  ProductsArr:any[];
  DepartmentsArr:any[];
  DivisionArr:any[];
  TagsArr:any[];
  selectedTab:string;
  ProductSegmentFilter:any[];

  IsStartPage:boolean;
  CurrentlyOpened:string;
  SearchQuery: string;
  IsResultsFound: boolean;
  OneDriveSearchResults:any[];
  SPOSearchResults:any[];
  EventsSearchResults:any[];
  MessageSearchResults:any[];
  
}  

let ProductsArr = [];
let DepartmentsArr = [];
let DivisionArr=[];
let TagsArr=[];
let GroupsArr=[];
let ProductsSegmentationArr = [];
let AllProductsArr = [];
let ProductVariantArr=[];
let ProductCategoryArr=[];
//var Globqueryarr = [];

let rowsPerPage;
let rows;
let rowsCount;
let pageCount; // avoid decimals
let numbers;

const NewWeb = Web("https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/"); 
export default class TaqeefDefinitions extends React.Component<ITaqeefDefinitionsProps,ITaqeefDefinitionsState,{}> {    
  public myRef;
  public constructor(props) {
    super(props);    
    this.myRef = React.createRef();   
    this.filterClear = this.filterClear.bind(this);    
    this.state = {
      value: [],
      Data:[],
      Product:[],
      ProductFilter:[],
      isClearable: true,   
      isSearchable:true,
      Productvariant:[],
      ProductCategory:[],
      CurrentlySelectedSearchTab:"All",
      ProductsArr:[],
      DepartmentsArr:[],
      DivisionArr:[],
      TagsArr:[],
      ProductSelected:"",
      DepartmentSelected:"",
      DivisionSelected:"",
      TagsSelected:"",
      selectedTab:"Definition",
      ProductSegmentFilter:[],

      IsStartPage:true,
      CurrentlyOpened:"",
      SearchQuery:"",
      IsResultsFound: true,
      OneDriveSearchResults:[],
      SPOSearchResults:[],
      EventsSearchResults:[],
      MessageSearchResults:[]
      
    }
  };
  

  public componentDidMount(){
    $(".close-icon").hide();
    $(".Filter_button").hide();
    $(".margin-top").hide(); 
    this.getDefinitionsMaster();
    this.getproductsegmentationfilter();
    // $(".def-Product").hide();
    $(".def-Department").hide();
    $(".def-Division").hide();
    $(".def-Tag").hide();

    $(".Product-Segment").hide();
    $(".Product-Variant").hide();
    $(".Product-Category").hide();

    
    $(".icon-bar a").on("click", function(){
      $(this).siblings().removeClass("active");
      $(this).addClass("active");
    });

    if(this.state.CurrentlySelectedSearchTab == "Definition"){
      $(".icon-bar a").siblings().removeClass("active");
      $(".def-class").addClass("active");
    }else if(this.state.CurrentlySelectedSearchTab == "ProductSegmentation"){
      $(".icon-bar a").siblings().removeClass("active");
      $(".prod-class").addClass("active");
    }else if(this.state.CurrentlySelectedSearchTab == "All"){
      if($.trim($("#txt-Search").val()) != ""){
        $(".icon-bar a").siblings().removeClass("active");
        $(".all-search").addClass("active");
      }
    }    
  }    

  public GetAllProducts(){
    NewWeb.lists.getByTitle("ProductSegmentation").items.select("Id","Group","Description","ProductImage","ProductType","ProductVariants","Category").get()
    .then((items)=>{
       if(items.length!=0){
        this.setState({Data:items});

        setTimeout(() => {
          this.pagination();
        }, 500);
       }      
    });
  }

  public GetAllDefinitions(){
    NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Product","Department","Division","Tags").top(4000).get()
    .then((items)=>{
      if(items.length!=0){
        this.setState({value:items});

        setTimeout(() => {
          this.pagination();
        }, 1200);
      }     
    });
  }
  
  public getDefinitionsMaster(){   
    var handler = this;  
    ProductsArr = [];
    DepartmentsArr =[];
    DivisionArr=[];
    TagsArr=[];
    ProductsSegmentationArr=[];
    $.ajax({
      url: "https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/_api/web/lists/GetByTitle('DefinitionsMaster')/fields?$filter=EntityPropertyName eq 'Department' or EntityPropertyName eq 'Division'",
      type: "GET",
      headers: {
        "accept": "application/json;odata=verbose",
      },
      success: function (data) {    
        for(var j = 0; j < data.d.results.length; j++){ 
          /*if(data.d.results[j].InternalName == "Product"){                   
            for(var i = 0; i < data.d.results[j].Choices.results.length; i++){ 
              //ProductsArr.push({ value: ''+data.d.results[j].Choices.results[i]+'', label: ''+data.d.results[j].Choices.results[i]+''});
              ProductsArr.push(data.d.results[j].Choices.results[i]);
              //AllProductsArr.push({value:''+data.d.results[j].Choices.results[i]+'',label:''+data.d.results[j].Choices.results[i]+''});
            }                    
          }*/
          if(data.d.results[j].InternalName == "Department"){                   
            for(var k = 0; k < data.d.results[j].Choices.results.length; k++){
              //DepartmentsArr.push({ value: ''+data.d.results[j].Choices.results[k]+'', label: ''+data.d.results[j].Choices.results[k]+''});
              DepartmentsArr.push(data.d.results[j].Choices.results[k]);
            }                   
          }
          else if(data.d.results[j].InternalName == "Division"){                   
            for(var m = 0; m < data.d.results[j].Choices.results.length; m++){
              //DivisionArr.push({ value: ''+data.d.results[j].Choices.results[m]+'', label: ''+data.d.results[j].Choices.results[m]+''});
              DivisionArr.push(data.d.results[j].Choices.results[m]);
            }                    
          }
          /*else if(data.d.results[j].InternalName == "Tags"){                   
            for(var n = 0; n < data.d.results[j].Choices.results.length; n++){
              //TagsArr.push({ value: ''+data.d.results[j].Choices.results[n]+'', label: ''+data.d.results[j].Choices.results[n]+''});
              TagsArr.push(data.d.results[j].Choices.results[n]);
            }
          }*/
        }     
        handler.setState({        
          DepartmentsArr: DepartmentsArr,
          DivisionArr: DivisionArr,          
        });
        //ProductsArr: ProductsArr,       
        //TagsArr: TagsArr
      },
      error: function (error) {
        console.log(JSON.stringify(error));
      }
    });    
  }

  

  public getEnteredQueryString(){  
  
        // $(".def-Product").hide();
        $(".def-Department").hide();
        $(".def-Division").hide();
        $(".def-Tag").hide();
        $(".Product-Segment").hide();
        $(".Product-Variant").hide();
        $(".Product-Category").hide();

        $(".Filter_button").hide();
      
        $("#no-result").hide()
    var input=$.trim($("#txt-Search").val());
    if(input.length < 3){
      $("#no-result").hide()
      this.setState({value:[],Data:[]});
      $(".margin-top").hide(); 
      $("#Search-err").show();
    }
    else{      
      if(this.Validation()){
        $("#Search-err").hide();
        this.setState({CurrentlySelectedSearchTab:"All"});
        $(".Filter_button").hide();
        $(".Def-Department").hide();
        $(".Def-Division").hide();
        $(".Def-Tags").hide();
        $(".Product-Category").hide();
        $(".Product-Variant").hide();
        $(".prod-segment").hide();
        $(".Definitionprod-segment, .prod-segment").hide();
        $(".margin-top").show(); 
        $(".all-products").show();
        var QueryString = $.trim($("#txt-Search").val());
        this.GetQueryResults(QueryString);
      }
    }
  }

  public GetQueryResults(QueryString){
    NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Product","Department","Division").
    filter("substringof('" + QueryString + "',Term) or substringof('" +QueryString+ "',Product) or substringof('" +QueryString+ "',Division) or substringof('" +QueryString+ "',Department)").get()
    .then((items)=>{
      if(items.length==0){
        $(".margin-top").hide();
        $("#no-result").show();
      }else{
        $("#no-result").hide();
        this.setState({value:items});
        setTimeout(() => {
          this.pagination();
        }, 700);
      }
      
    }); 

    NewWeb.lists.getByTitle("ProductSegmentation").items.select("Id","Group","Description","ProductImage","ProductType","ProductVariants","Category").
    filter("substringof('" + QueryString + "',Group) or substringof('" +QueryString+ "',ProductType) or substringof('" +QueryString+ "',ProductVariants) or substringof('" +QueryString+ "',Category) " ).get()
    .then((items)=>{
      // if(items.length==0){
      //   $("#no-result").show();
      // }else{
      //   $("#no-result").hide();
      // }
      this.setState({Data:items});

      setTimeout(() => {
        this.pagination();
      }, 700);
    }); 
  }

  // Search Result
  public SearchResult(e){     
    
    var input=$.trim($("#txt-Search").val());
    var handler = this;
    if(e.keyCode == 13 && this.state.CurrentlySelectedSearchTab=="All"){  
      //$(".Filter_button").hide();
      if(input.length < 3){
        $("#no-result").hide()
        $("#Search-err").show();
        $(".margin-top").hide();
        this.setState({value:[],Data:[]});
      }
      else{
        $("#Search-err").hide();        
        this.setState({CurrentlySelectedSearchTab:"All"});
        if(this.state.CurrentlySelectedSearchTab == "All"){
          $(".icon-bar a").siblings().removeClass("active");
          $(".all-search").addClass("active");
        }
        handler.getEnteredQueryString();
        $(".margin-top").show();   
      }     
    }
    else if(e.keyCode == 13 && this.state.CurrentlySelectedSearchTab=="Definition"){  
     var product=$(".def-Product").val();
     var Department=$(".def-Department").val();
     var Division=$(".def-Division").val();
     var Tag=$(".def-Tag").val();
if(product==null || Department==null || Division==null || Tag==null){
this.DefinitionMasterSerach();
}else if(input.length < 3){
        $("#no-result").hide()
        $("#Search-err").show();
        $(".margin-top").hide();
        this.setState({value:[],Data:[]});
      }
      else {
      $("#Search-err").hide();   
      
      $(".icon-bar a").siblings().removeClass("active");
      $(".def-class").addClass("active");
      this.DefinitionSearchAfterFilter();
    }
    }else if(e.keyCode == 13 && this.state.CurrentlySelectedSearchTab=="ProductSegmentation"){
      var search=$.trim($("#txt-Search").val());
      if(search==''){
      if(input.length < 3){
       
        $("#no-result").hide()
        $("#Search-err").show();
        $(".margin-top").hide();
        this.setState({value:[],Data:[]});
      }
      else{
        $("#Search-err").hide();
      if($("#def-Product").length > 1) {
        $(".Filter_button").show();
      }
      $(".icon-bar a").siblings().removeClass("active");
      $(".prod-class").addClass("active");
      this.ProductSearchAfterFilter();
    }
  }
  else if(search!=''){
    this.SearchProductOnly();
  }
  }
  
  
}
  //Product FIlter
  public getproductsegmentationfilter(){
    GroupsArr = [];
    ProductCategoryArr = [];
    ProductVariantArr = [];
    var handler=this;
    $.ajax({
      url: "https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/_api/web/lists/GetByTitle('ProductSegmentation')/fields?$filter=EntityPropertyName eq 'Group' or EntityPropertyName eq 'Category' or EntityPropertyName eq 'ProductVariants'",
      type: "GET",
      headers: {
        "accept": "application/json;odata=verbose",
      },
      success: function (data) {             
        for(var j = 0; j < data.d.results.length; j++){ 
          if(data.d.results[j].InternalName == "Group"){                   
            for(var i = 0; i < data.d.results[j].Choices.results.length; i++){ 
              GroupsArr.push(data.d.results[j].Choices.results[i]);
            }                    
          }
          else if(data.d.results[j].InternalName == "Category"){                   
            for(var k = 0; k < data.d.results[j].Choices.results.length; k++){
              ProductCategoryArr.push(data.d.results[j].Choices.results[k]);
            }                   
          }
          else if(data.d.results[j].InternalName == "ProductVariants"){                   
            for(var m = 0; m < data.d.results[j].Choices.results.length; m++){
              ProductVariantArr.push(data.d.results[j].Choices.results[m]);
            }                    
          }
        }
        handler.setState({
          ProductSegmentFilter: GroupsArr, 
          Productvariant: ProductVariantArr,
          ProductCategory: ProductCategoryArr
        });          
      },
      error: function (error) {
        console.log(JSON.stringify(error));
      }
    });
  }


  public FilterBasedonAllProduct(SelcectedProduct){
    var QueryString = $.trim($("#txt-Search").val());
    NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Product","Department","Division").
    filter("Product eq '"+SelcectedProduct+"' and substringof('" + QueryString + "',Term) or substringof('" +QueryString+ "',Division) or substringof('" +QueryString+ "',Department)").get()       
    .then((items)=>{
    
      this.setState({value:items});

      setTimeout(() => {
        this.pagination();
      }, 700);
    });
    
    NewWeb.lists.getByTitle("ProductSegmentation").items.select("Id","Group","Description","ProductImage","ProductType","ProductVariants","Category").
    filter("Group eq '"+SelcectedProduct+"' and substringof('" +QueryString+ "',ProductType) or substringof('" +QueryString+ "',ProductVariants) or substringof('" +QueryString+ "',Category) " ).get()
    .then((items)=>{
      if(items.length==0){
        $(".margin-top").hide();
        $("#no-result").show();
      }else{
        $("#no-result").hide();
        this.setState({Data:items});
        setTimeout(() => {
          this.pagination();
        }, 700);
      }
      
    });
  }

 public DefinitionMasterSerach(){  
  
  $("#numbers").empty();

   this.setState({Data:[]})
    $(".Product-Segment").hide();
    $(".Product-Variant").hide();
    $(".Product-Category").hide();

    $(".def-Product").show();
    $(".def-Department").show();
    $(".def-Division").show();
    $(".def-Tag").show();

  var input=$.trim($("#txt-Search").val());
  if(input.length == 0){
    $(".def-Product").show();
    $(".def-Department").show();
    $(".def-Division").show();
    $(".def-Tag").show();
    this.GetAllDefinitions();
  }
  if(input.length >= 1 && input.length < 3){
    $("#no-result").hide()
    $("#Search-err").show();
    $(".margin-top").hide();
    this.setState({value:[],Data:[]});
  }
  else{
    $("#Search-err").hide();
    $("#no-result").hide()
    $(".Product-Segment").val('');
    $(".Product-Variant").val('');
    $(".Product-Category").val('');
  
   this.setState({CurrentlySelectedSearchTab:"Definition"});
  var QueryString = $.trim($("#txt-Search").val());
  if(QueryString != ""){
    this.SearchDefinitionOnly();
  }else{
    $(".Filter_button").hide();


    $(".def-Product").show();
    $(".def-Department").show();
    $(".def-Division").show();
    $(".def-Tag").show();

    $(".Product-Segment").hide();
    $(".Product-Variant").hide();
    $(".Product-Category").hide();
  }
} 
}
  

  public SearchDefinitionOnly(){     
    $(".Filter_button").hide();


    $(".def-Product").show();
    $(".def-Department").show();
    $(".def-Division").show();
    $(".def-Tag").show();

    $(".Product-Segment").hide();
    $(".Product-Variant").hide();
    $(".Product-Category").hide();
 
      var tab=this.state.selectedTab;
      this.setState({CurrentlySelectedSearchTab:"Definition"});
   
      
     
      $(".prod-segment").hide();
      $(".Definitionprod-segment").show();
      $(".all-products").hide();
      this.setState({Data:[]});
      var QueryString = $.trim($("#txt-Search").val());
      NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Product","Department","Division").
      filter("substringof('" + QueryString + "',Term) or substringof('" +QueryString+ "',Product) or substringof('" +QueryString+ "',Division) or substringof('" +QueryString+ "',Department)").get()
      .then((items)=>{
        if(items.length==0){
          $(".margin-top").hide();
          $("#no-result").show();
        }else{
          $(".margin-top").show();
          $("#no-result").hide();
          this.setState({value:items});
          setTimeout(() => {
            this.pagination();
          }, 700);
        }
        
      });
    // }
  }

  public ProductSearchMaster(){ 
    $("#numbers").empty();
    this.setState({value:[]})
      $(".def-Product").hide();
      $(".def-Department").hide();
      $(".def-Division").hide();
      $(".def-Tag").hide();

      $(".Product-Segment").show();
      $(".Product-Variant").show();
      $(".Product-Category").show();
    var input=$.trim($("#txt-Search").val());
    if(input.length == 0){
      $(".Product-Segment").show();
      $(".Product-Variant").show();
      $(".Product-Category").show();
      this.GetAllProducts();
    }
    if(input.length >= 1 && input.length < 3){
      $("#no-result").hide()
      $("#Search-err").show();
      $(".margin-top").hide();
      this.setState({value:[],Data:[]});
    }
    else{
      $("#Search-err").hide();
      $("#no-result").hide()  
      $(".def-Product").val('');
      $(".def-Department").val('');
      $(".def-Division").val('');
      $(".def-Tag").val('');

    this.setState({CurrentlySelectedSearchTab:"ProductSegmentation"});
    var QueryString = $.trim($("#txt-Search").val());
    if(QueryString != ""){
      this.SearchProductOnly();
    }else{
      $(".def-Product").hide();
      $(".def-Department").hide();
      $(".def-Division").hide();
      $(".def-Tag").hide();

      $(".Product-Segment").show();
      $(".Product-Variant").show();
      $(".Product-Category").show();
    }
  }
}

  public SearchProductOnly(){  
    var QueryString = $.trim($("#txt-Search").val());
    if(QueryString.length < 3){
      $("#Search-err").show();
    }
    else{
    $("#Search-err").hide();
    $(".margin-top").show();     
    $(".def-Product").hide();
    $(".def-Department").hide();
    $(".def-Division").hide();
    $(".def-Tag").hide();

    $(".Product-Segment").show();
    $(".Product-Variant").show();
    $(".Product-Category").show();
 

      this.setState({CurrentlySelectedSearchTab:"ProductSegmentation"});
      
      this.setState({value:[]});
      
      NewWeb.lists.getByTitle("ProductSegmentation").items.select("Id","Group","Description","ProductImage","ProductType","ProductVariants","Category").
      filter("substringof('" + QueryString + "',Group) or substringof('" +QueryString+ "',ProductType) or substringof('" +QueryString+ "',ProductVariants) or substringof('" +QueryString+ "',Category) " ).get()
      .then((items)=>{
        if(items.length==0){
          $(".margin-top").hide();
          $("#no-result").show();
        }else{
          $("#no-result").hide();
          this.setState({Data:items});
          setTimeout(() => {
            this.pagination();
          }, 700);
        }
        
      });    
     }
  }

  public Validation(){
    let formstatus = true;
    var SearchQuery = $.trim($("#txt-Search").val());
    if(formstatus == true && SearchQuery != ''){
      $("#txt-err-msg-search").hide();
      return formstatus;      
    }else{
      $("#txt-err-msg-search").show();
      formstatus = false;      
    }
    return formstatus;
  }


  
  public filterClear(){
    $(".Filter_button").hide();
    $(".def-Product").val('');
    $(".def-Department").val('');
    $(".def-Division").val('');
    $(".def-Tag").val('');

    $(".Product-Segment").val('');
    $(".Product-Variant").val('');
    $(".Product-Category").val('');

    if(this.state.CurrentlySelectedSearchTab=="Definition"){
      this.DefinitionMasterSerach();
    }
    else if(this.state.CurrentlySelectedSearchTab=="ProductSegmentation"){
    this.ProductSearchMaster();
    }

  }

  public ClearSearchInput(){
    $("#txt-Search").val('');
  }

  public iconRemove(){
    if($("#txt-Search").length==0){
      $(".close-icon").hide();
     }
     else if($("#txt-Search").length > 0){
      $(".close-icon").show();
     } 
  }


  public masterdefinitionfilter(){
    $("#Search-err").hide();
    $(".margin-top").show(); 
    if($("#def-Division").length == 1 || $("#def-Department").length == 1) {
      $(".Filter_button").show();
    }
    var query=$.trim($("#txt-Search").val());
    if(query==''){
      this.DefinitionSearchBeforeFilter();
    }
    else if(query!=''){
      this.DefinitionSearchAfterFilter();
    }
  }

  public masterProductionfilter(){
    $("#Search-err").hide();
    $(".margin-top").show(); 
    if($("#def-Division").length == 1 || $("#def-Department").length == 1) {
      $(".Filter_button").show();
    }
    var query=$.trim($("#txt-Search").val());
    if(query==''){
      this.ProductSearchBeforeFilter();
    }
    else if(query!=''){
      this.ProductSearchAfterFilter();
    }
  }


  public DefinitionSearchAfterFilter(){  
    debugger;
    var SearAftFil = [];
    var queryarr = [];  
    var filterquery;
    var QueryString = $.trim($("#txt-Search").val()); 
    var DefDepartment=$("#def-Department").val();
    var DefDivision=$("#def-Division").val();
    //var DefTag=$("#def-Tag").val();
     if(DefDepartment != "" && DefDepartment != 'undefined' && DefDepartment != null){
       //filterquery = `Department eq ${DefDepartment}`;
       queryarr.push(`Department eq '${DefDepartment}' and `);
     }
     if(DefDivision != "" && DefDivision != 'undefined' && DefDivision != null){
       //filterquery = `Division eq ${DefDivision}`;
       queryarr.push(`Division eq '${DefDivision}' and `);
     }
     /*if(DefTag != "" && DefTag != 'undefined' && DefTag != null){
       //filterquery = `Tags eq ${DefTag}`;
       queryarr.push(`Tags eq '${DefTag}' and `);
     }*/
     filterquery = queryarr.join("");

     if(DefDivision != null && DefDepartment != null){
      this.setState({value:[]});      
      NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Department","Division").
      filter(""+filterquery+" substringof('" + QueryString + "',Term) or substringof('" +QueryString+ "',Department) or substringof('" +QueryString+ "',Division)").get()// and substringof('" +QueryString+ "',Department) or substringof('" +QueryString+ "',Division)
      .then((items)=>{
        
        if(items.length==0){
          $(".margin-top").hide();
          $("#no-result").show();
          this.setState({value:[]});
          $("#numbers").empty();
        }else{        
          $(".margin-top").show();
            $("#no-result").hide();
            this.setState({value:items}); 
            setTimeout(() => {
              this.pagination();
            }, 700);        
        }
      });
     }

    if(DefDivision != null && DefDepartment == null){
      this.setState({value:[]});      
      NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Department","Division").
    filter(""+filterquery+" substringof('" + QueryString + "',Term) or substringof('" +QueryString+ "',Division)").get()// and substringof('" +QueryString+ "',Department) or substringof('" +QueryString+ "',Division)
    .then((items)=>{
      
      if(items.length==0){
        $(".margin-top").hide();
        $("#no-result").show();
        this.setState({value:[]});
        $("#numbers").empty();
      }else{        
        $(".margin-top").show();
          $("#no-result").hide();
          this.setState({value:items}); 
          setTimeout(() => {
            this.pagination();
          }, 700);        
      }
    });
    }
    if(DefDepartment != null && DefDivision == null){
      this.setState({value:[]});      
      NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Department","Division").
    filter(""+filterquery+" substringof('" + QueryString + "',Term) or substringof('" +QueryString+ "',Department)").get()// and substringof('" +QueryString+ "',Department) or substringof('" +QueryString+ "',Division)
    .then((items)=>{
      
      if(items.length==0){
        $(".margin-top").hide();
        $("#no-result").show();
        this.setState({value:[]});
        $("#numbers").empty();
      }else{        
        $(".margin-top").show();
          $("#no-result").hide();
          this.setState({value:items}); 
          setTimeout(() => {
            this.pagination();
          }, 700);        
      }
    });
    }

    
    
    
  }

  public ProductSearchAfterFilter(){   
    var queryarr = []; 
    var QueryString = $.trim($("#txt-Search").val());
    var Category=$("#Product-Category").val();
    var ProVariant=$("#Product-Variant").val();
    var ProSegment=$("#Product-Segment").val();

    if(ProSegment != "" && ProSegment != 'undefined' && ProSegment != null){      
      queryarr.push(`Group eq '${ProSegment}' and `);
    }
    if(ProVariant != "" && ProVariant != 'undefined' && ProVariant != null){      
      queryarr.push(`ProductVariants eq '${ProVariant}' and `);
    }  
    if(Category != "" && Category != 'undefined' && Category != null){    
      queryarr.push(`Category eq '${Category}' and `);
    }     
    var filterquery = queryarr.join("");    

    NewWeb.lists.getByTitle("ProductSegmentation").items.select("Group","Description","ProductImage","ProductType","ProductVariants","Category").    
    filter(""+filterquery+" substringof('" +QueryString+ "',ProductType) or substringof('" +QueryString+ "',ProductVariants) or substringof('" +QueryString+ "',Category) " ).get()
    .then((items)=>{
      if(items.length==0){
        $(".margin-top").hide();
        $("#no-result").show();
        this.setState({Data:[]});
        $("#numbers").empty();
      }else{
        $("#no-result").hide();
        this.setState({Data:items});
        setTimeout(() => {
          this.pagination();
        }, 700);
      }
      
    });
  }



  public DefinitionSearchBeforeFilter(){ 

    var Globqueryarr = [];
    var filterquery;
    var DefDepartment=$("#def-Department").val();
    var DefDivision=$("#def-Division").val();
    //var DefTag=$("#def-Tag").val();
    if(DefDepartment != "" && DefDepartment != 'undefined' && DefDepartment != null){
      Globqueryarr.push(` Department eq '${DefDepartment}' `);
     }
     if(DefDivision != "" && DefDivision != 'undefined' && DefDivision != null){
      Globqueryarr.push(` Division eq '${DefDivision}' `);
     }
     /*if(DefTag != "" && DefTag != 'undefined' && DefTag != null){
      Globqueryarr.push(` Tags eq '${DefTag}' `);
     }*/
     filterquery = Globqueryarr.join(" and ");
    NewWeb.lists.getByTitle("DefinitionsMaster").items.select("Term","Description","Department","Division").
    filter(""+filterquery+"").get()
    .then((items)=>{
       
      if(items.length==0){
        $(".margin-top").hide();
        $("#no-result").show();
        this.setState({value:[]});
        $("#numbers").empty();
      }else{
        $(".margin-top").show();
        $("#no-result").hide();
        this.setState({value:items});
        setTimeout(() => {
          this.pagination();
        }, 700);
      }
    });
  }

  public ProductSearchBeforeFilter(){   
    var GlobProqueryarr = []; 
    var filterquery;
    var Category=$("#Product-Category").val();
    var ProVariant=$("#Product-Variant").val();
    var ProSegment=$("#Product-Segment").val();

    if(ProSegment != "" && ProSegment != 'undefined' && ProSegment != null){      
      GlobProqueryarr.push(` Group eq '${ProSegment}'`);
    }
    if(ProVariant != "" && ProVariant != 'undefined' && ProVariant != null){      
      GlobProqueryarr.push(` ProductVariants eq '${ProVariant}'`);
    }  
    if(Category != "" && Category != 'undefined' && Category != null){    
      GlobProqueryarr.push(` Category eq '${Category}'`);
    }     
    filterquery = GlobProqueryarr.join(" and ");   

    NewWeb.lists.getByTitle("ProductSegmentation").items.select("Id","Group","Description","ProductImage","ProductType","ProductVariants","Category").    
    filter(""+filterquery+"").get()
    .then((items)=>{
      if(items.length==0){
        $(".margin-top").hide();
        $("#no-result").show();
        this.setState({Data:[]});
        $("#numbers").empty();
      }else{
        $("#no-result").hide();
        this.setState({Data:items});
        setTimeout(() => {
          this.pagination();
        }, 700);
      }
      
    });
  }

  //unified search///////////////////////////////////////


  public ValidateSearchInputField(){
    let txtsearchinput = $('#txt-Search').val(); 
    let status = true;   
    if(status == true && txtsearchinput != ''){
      $('.error-input').hide();  
      return status;    
    }else{
      $('.error-input').show();
      $('.txt-search-input').focus();
      status = false;
    }
    return status;
  }


  public SearchDrive = (searchtext) => {
    if(this.ValidateSearchInputField()){
      // $(".Loader").addClass('open');
      this.setState({IsStartPage:false});
      this.setState({CurrentlyOpened:"Drive"});
      // this.ChnageActiveClass();
      // this.ClearResultsbeforeNew();
      this.setState({SearchQuery:""+searchtext+""});
      var searchquery = `/me/drive/root/search(q='${searchtext}')`;
      const DriveSearchitems = this.props.graphClient.api(''+searchquery+'').version('v1.0').get((err: any, response: any): void => {
        if(response.value.length != 0){
          this.setState({ 
            IsResultsFound: true,           
            OneDriveSearchResults: response.value
          }); 
        }else{
          this.setState({
            IsResultsFound: false
          });
          $(".Loader").removeClass('open');
        }                              
      });       
    }
  }



  public SearchSharePoint = (searchtext) => {    
    if(this.ValidateSearchInputField()){      
      $(".Loader").addClass('open');   
      this.setState({IsStartPage:false});
      this.setState({CurrentlyOpened:"SharePoint"});
      // this.ChnageActiveClass();
      // this.ClearResultsbeforeNew();
      this.setState({SearchQuery:""+searchtext+""});    
      try {
          const SharePointSearch = this.props.graphClient.api('https://graph.microsoft.com/v1.0/search/query').version('v1.0').post(
            {
              "requests": [
                  {
                      "entityTypes": [
                          "listItem"
                      ],
                      "query": {
                          "queryString": `${searchtext}`
                      }
                  }
              ]
          }
          ).then((response) => {   
            if(response.value[0].hitsContainers[0].total != 0){           
              this.setState({    
                IsResultsFound: true,            
                SPOSearchResults: response.value[0].hitsContainers[0].hits               
              });
            }else{    
              $(".Loader").removeClass('open');        
              this.setState({
                IsResultsFound: false
              });              
            }
          });
      }
      catch (error) {
          console.log(error);
      }      
    }
  }


  public SearchEvents = (searchtext) => {      
    if(this.ValidateSearchInputField()){
      $(".Loader").addClass('open');
      this.setState({IsStartPage:false});
      this.setState({CurrentlyOpened:"Calendar"});
      // this.ChnageActiveClass();
      // this.ClearResultsbeforeNew(); 
      this.setState({SearchQuery:""+searchtext+""});    
      try {
          const EventSearch = this.props.graphClient.api('https://graph.microsoft.com/v1.0/search/query').version('v1.0').post(
            {
              "requests": [
                  {
                      "entityTypes": [
                          "event"
                      ],
                      "query": {
                          "queryString": `${searchtext}`
                      }
                  }
              ]
          }
          ).then((response) => {                         
            if(response.value[0].hitsContainers[0].total != 0){
              this.setState({   
                IsResultsFound: true,             
                EventsSearchResults: response.value[0].hitsContainers[0].hits
              });
            }else{
              $(".Loader").removeClass('open');
              this.setState({
              IsResultsFound: false
              });  
            }                        
          });
      }
      catch (error) {
          console.log(error);
      }                                    
    }
  }



  public SearchMessages = (searchtext) => {      
    if(this.ValidateSearchInputField()){
      $(".Loader").addClass('open');
      this.setState({IsStartPage:false});
      this.setState({CurrentlyOpened:"EMail"});
      // this.ChnageActiveClass();
      // this.ClearResultsbeforeNew();     
      this.setState({SearchQuery:""+searchtext+""});
      try {
          const MessageSearch = this.props.graphClient.api('https://graph.microsoft.com/v1.0/search/query').version('v1.0').post(
            {
              "requests": [
                  {
                      "entityTypes": [
                          "message"
                      ],
                      "query": {
                          "queryString": `${searchtext}`
                      }
                  }
              ]
          }
          ).then((response) => {   
            if(response.value[0].hitsContainers[0].total != 0){
              this.setState({      
                IsResultsFound: true,          
                MessageSearchResults: response.value[0].hitsContainers[0].hits
              });
            }else{    
              $(".Loader").removeClass('open');        
              this.setState({
                IsResultsFound: false
              });              
            }
          });
      }
      catch (error) {
          console.log(error);
      }      
    }
  }

  public pagination(){
    $("#numbers").empty();
    rowsPerPage = 25;
    rows = $('.Pagination-element-wrap div.search-results');
    rowsCount = rows.length;
    pageCount = Math.ceil(rowsCount / rowsPerPage); // avoid decimals
    numbers = $('#numbers');
    
    // Generate the pagination.
    for (var i = 0; i < pageCount; i++) {   
      if(i == 0)   
      numbers.append('<li className="page-item active"><a className="page-link no-border" href="#">' + (i+1) + '</a></li>')
      else
      numbers.append('<li className="page-item"><a className="page-link no-border" href="#">' + (i+1) + '</a></li>')
    }
      
    // Mark the first page link as active.
    $('#numbers li:first-child a').addClass('active');

    // Display the first set of rows.
    displayRows(1);
    
    // On pagination click.
    $('#numbers li a').on("click",function(e) {
      
      var $this = $(this);
      
      e.preventDefault();
      
      // Remove the active class from the links.
      $('#numbers li a').removeClass('active');
      
      // Add the active class to the current link.
      $this.addClass('active');
      
      // Show the rows corresponding to the clicked page ID.
      
      displayRows($this.text());
    });
  }

public render(): React.ReactElement<ITaqeefDefinitionsProps> {     
  const SearchResults: JSX.Element[] = this.state.value.map(function (item, key) { 
    return (     
      <div className="search-results">
        <ul>      
          <li><h4>{item.Term}</h4></li>
          <li><h5>{item.Description}</h5></li>
        </ul>   
        <div>
          <div className="product-hover">
            <p className="department-value">{item.Department}</p>
            <span className="product-hovertext">Department</span>
          </div>
          <div className="product-hover">
            <p className="division-value">{item.Division}</p>
            <span className="product-hovertext">Division</span>
          </div>
        </div>
      </div>    
    );        
  });
  const ProductResults: JSX.Element[] = this.state.Data.map(function (item, key) {  
    if(item.Category != "" && item.Category != null && item.Category != 'undefined'){
    for(var i = 0; i <item.Category.length; ){          
      if(item.Category[i]=="General"){              
        var Genralbrand="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/brand-1.png";
        setTimeout(() => {
          if ($(`.generalimg-clas-${key}`)[0]){
            // Do something if class exists
          } else {
            $("#"+item.Id+"-Brand-availability").append(`<img class="brand-img generalimg-clas-${key}" src=${Genralbrand} alt="generallogo" title="General"></img>`);
          }          
        }, 100);              
      }
      if(item.Category[i]=="Midea"){             
        var Mideabrand="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/brand-2.png";
        setTimeout(() => {
          if ($(`.mideaimg-clas-${key}`)[0]){
            // Do something if class exists
          } else {
            $("#"+item.Id+"-Brand-availability").append(`<img class="brand-img mideaimg-clas-${key}" src=${Mideabrand} alt="midealogo" title="Midea"></img>`);
          }          
        }, 100);          
      }
      if(item.Category[i]=="Novair"){          
        var Clivetbrand="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/NovAir.png";
        setTimeout(() => {
          if ($(`.Clivetimg-clas-${key}`)[0]){
            // Do something if class exists
          } else {
            $("#"+item.Id+"-Brand-availability").append(`<img class="brand-img Clivetimg-clas-${key}" src=${Clivetbrand} alt="novairlogo" title="Novair"></img>`);
          }          
        }, 100);          
      }
      if(item.Category[i]=="Clint"){          
        var Clivetbrand="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/clint.png";
        setTimeout(() => {
          if ($(`.Clint-clas-${key}`)[0]){
            // Do something if class exists
          } else {
            $("#"+item.Id+"-Brand-availability").append(`<img class="brand-img Clint-clas-${key}" src=${Clivetbrand} alt="clintlogo" title="Clint"></img>`);
          }          
        }, 100);          
      }
      if(item.Category[i]=="Trosten"){              
        var Trostenbrand="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/trosten.png";
        setTimeout(() => {
          if ($(`.Trostenimg-clas-${key}`)[0]){
            // Do something if class exists
          } else {
            $("#"+item.Id+"-Brand-availability").append(`<img class="brand-img Trostenimg-clas-${key}" src=${Trostenbrand} alt="trostenlogo" title="Trosten"></img>`);
          }          
        }, 100);
      }
      i++;
    }

    let RawImageTxt = item.ProductImage;
      if(RawImageTxt != "" && RawImageTxt != null){
      var ImgObj = JSON.parse(RawImageTxt);
      var Category = item.Category;            
      return (
        <div className="search-results clearfix">
          <div className="prod-img-wrap">
          <img className="img" src={`${ImgObj.serverRelativeUrl}`} alt="Window AC"  style={{width:"140px",marginRight:"15px"}} height="90px"/>
          </div>
          <div className="prod-img-content">
          <div className="content-head">
            <p className="sub-head color-blue ">{item.ProductType}</p>
            <p className="font-12 color-gray ">{item.Description}</p>
            <ul  id={item.Id+"-Brand-availability"}> 
            </ul>
            <div className="Product-variant-wrap">
              {item.ProductVariants && item.ProductVariants.map((item,key)=>{
                return(              
                <div className="product-hover">
                  <p className="department-value">{item}</p>
                  <span className="product-hovertext">Variant</span>
                </div>                            
                );
              })}   
            </div>         
          </div> 
          </div>         
        </div>                
      );
    }    
  }else{
    let RawImageTxtt = item.ProductImage;
      if(RawImageTxtt != "" && RawImageTxtt != null){
      var ImgObjj = JSON.parse(RawImageTxtt);
    return(
      <div className="search-results clearfix">
          <div className="prod-img-wrap">
          <img className="img" src={`${ImgObjj.serverRelativeUrl}`} alt="Window AC"  style={{width:"140px",marginRight:"15px"}} height="90px"/>
          </div>
          <div className="prod-img-content">
          <div className="content-head">
            <p className="sub-head color-blue ">{item.ProductType}</p>
            <p className="font-12 color-gray ">{item.Description}</p>
            <ul  id={item.Id+"-Brand-availability"}> 
            </ul>                     
          </div> 
          </div>         
        </div>
    );
      }
  }   
  });
  const GroupProductResults: JSX.Element[] = this.state.ProductFilter.map(function (item, key) {             
      return(
        <>
          <input type="checkbox" id={key+"Product"} name="Product" value={item}/>
          <label  className="font-12">{item}</label>
        </>
      );           
  });


  const DefinitionProductsOptions: JSX.Element[] = this.state.ProductsArr.map(function(item,key) {
    return(
      <option value={item}>{item}</option>
    );
  });

  const DefinitionDepartmentOptions: JSX.Element[] = this.state.DepartmentsArr.map(function(item,key) {
    return(
      <option value={item}>{item}</option>
    );
  });

  const DefinitionTagOptions: JSX.Element[] = this.state.TagsArr.map(function(item,key) {
    return(
      <option value={item}>{item}</option>
    );
  });
  const DefinitionDivisionOptions: JSX.Element[] = this.state.DivisionArr.map(function(item,key) {
    return(
      <option value={item}>{item}</option>
    );
  });

  const ProductSegmentOptions: JSX.Element[] = this.state.ProductSegmentFilter.map(function(item,key) {
    return(
      <option value={item}>{item}</option>
    );
  });

  const ProductVariantOptions: JSX.Element[] = this.state.Productvariant.map(function(item,key) {
    return(
      <option value={item}>{item}</option>
    );
  });

  const ProductCategoryOptions: JSX.Element[] = this.state.ProductCategory.map(function(item,key) {
    return(
      <option value={item}>{item}</option>
    );
  });


//unified result code /////////////////////////////////
const OneDriveSearchResults: JSX.Element[] = this.state.OneDriveSearchResults.map(function (item, key) {
  var handler=this;
  $(".Loader").removeClass('open');
  var filename = item.name.split(".");
  var extension = filename[filename.length - 1];
  var img = "";      

  if(item.folder){
    return(
      <li className="clearfix">
        <div className="search-img-show">
       
       
          <img src={`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/folder.png`} alt="image" />
         
        </div>
        <div className="search-contens-show">
          <a href={item.webUrl} target='_blank' data-interception='off'> {item.name} </a>
          <h4> {item.lastModifiedBy.user.displayName} modified on {moment(item.fileSystemInfo.lastModifiedDateTime).format("DD MMM YYYY")}  </h4>
          <p> </p>
        </div>
      </li>
    );
  }else{
    if(extension == 'docx' || extension == 'doc'){
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/WordFluent.png`;                
    }
    if(extension == 'pdf'){
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/PDF.JPG`;
    }
    if(extension == 'xlsx'){
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/ExcelFluent.png`;
    }
    if(extension == 'pptx'){
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/PPTFluent.png`;
    }
    if(extension == 'url'){
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/URL.png`;
    }
    if(extension == 'txt'){ 
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/font.png`;
    }
    if(extension == 'css' || extension == 'sppkg' || extension == 'ts' || extension == 'tsx' || extension == 'html' || extension == 'aspx' || extension == 'ts' || extension == 'js' || extension == 'map' || extension == 'php' || extension == 'json' || extension == 'xml'){
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/coding.png`;
    }
    if(extension == 'png' || extension == 'PNG' || extension == 'JPG' || extension == 'JPEG'  || extension == 'SVG' || extension == 'svg' || extension == 'jpg' || extension == 'jpeg' || extension == 'gif'){
      img = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/image.png`;
    }
    if(extension == "zip" || extension == "rar"){
      img=`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/zip.svg`;
    }
    return(
      <li className="clearfix">
        <div className="search-img-show">
          <img src={img} alt="image" />
        </div>
        <div className="search-contens-show">
          <a href={item.webUrl} target='_blank' data-interception='off'> {item.name} </a>
          <h4> {item.lastModifiedBy.user.displayName} modified on {moment(item.fileSystemInfo.lastModifiedDateTime).format("DD MMM YYYY")}  </h4>
          <p></p>
        </div>
      </li>
    );
  }
});



const SPOSearchResults: JSX.Element[] = this.state.SPOSearchResults.map(function (item, key) {   
  var handler=this;
  $(".Loader").removeClass('open');
  var DocName = "";
  var FileTypeImg = ""; 
if(item.resource.lastModifiedBy != undefined && item.resource.lastModifiedBy != null && item.resource.lastModifiedBy != "undefined" && item.resource.lastModifiedBy != "null"){         
  var fileextention = item.resource.webUrl.split(".");
  var filenamefromurl = item.resource.webUrl.split("/");
  var DocumentNamewithextention = filenamefromurl[filenamefromurl.length - 1];
  var DocNameExtract = DocumentNamewithextention.split(".");
  var extension = fileextention[fileextention.length - 1];
  
  var incStr = extension.includes("aspx?ID="); 
  if(incStr == true){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/common-icon.png`;
  }
  
  if(DocNameExtract.length == 1){
    DocName = DocNameExtract;
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/common-icon.png`;
  } else{
    DocName = DocNameExtract[DocNameExtract.length - 2];
  }

  if(extension == 'docx' || extension == 'doc'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/WordFluent.png`;                
  }
  if(extension == 'pdf'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/PDF.JPG`;
  }
  if(extension == 'xlsx'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/ExcelFluent.png`;
  }
  if(extension == 'pptx'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/PPTFluent.png`;
  }
  if(extension == 'url'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/URL.png`;
  }
  if(extension == 'txt'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/font.png`;
  }
  if(extension == 'css' || extension == 'sppkg' || extension == 'ts' || extension == 'tsx' || extension == 'html' || extension == 'aspx' || extension == 'ts' || extension == 'js' || extension == 'map' || extension == 'php' || extension == 'json' || extension == 'xml'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/coding.png`;
  }
  if(extension == 'png' || extension == 'PNG' || extension == 'JPG' || extension == 'JPEG'  || extension == 'SVG' || extension == 'svg' || extension == 'jpg' || extension == 'jpeg' || extension == 'gif'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/image.png`;
  }
  if(extension == "zip" || extension == "rar"){
    FileTypeImg=`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/zip.svg`;
  }    

  return(
    <li className="clearfix">
      <div className="search-img-show">
        <img src={FileTypeImg} alt="image" />
      </div>
      <div className="search-contens-show">
        <a href={item.resource.webUrl} target='_blank' data-interception='off'> {DocName} </a>
        <h4> {item.resource.lastModifiedBy.user.displayName} modified on {moment(item.resource.lastModifiedDateTime).format("DD MMM YYYY hh:mm a")}  </h4>  
        <p><Markup content={item.summary} /></p>            
      </div>
    </li>
  );  
}else{
  var fileextention = item.resource.webUrl.split(".");
  var filenamefromurl = item.resource.webUrl.split("/");
  var DocumentNamewithextention = filenamefromurl[filenamefromurl.length - 1];
  var DocNameExtract = DocumentNamewithextention.split(".");
  var extension = fileextention[fileextention.length - 1];
  
  var incStr = extension.includes("aspx?ID="); 
  if(incStr == true){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/common-icon.png`;
  }

  if(DocNameExtract.length == 1){
    DocName = DocNameExtract;
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/common-icon.png`;          
  } else{
    DocName = DocNameExtract[DocNameExtract.length - 2];
  }

  if(extension == 'docx' || extension == 'doc'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/WordFluent.png`;                
  }
  if(extension == 'pdf'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/PDF.JPG`;
  }
  if(extension == 'xlsx'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/ExcelFluent.png`;
  }
  if(extension == 'pptx'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/PPTFluent.png`;
  }
  if(extension == 'url'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/URL.png`;
  }
  if(extension == 'txt'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/font.png`;
  }
  if(extension == 'css' || extension == 'sppkg' || extension == 'ts' || extension == 'tsx' || extension == 'html' || extension == 'aspx' || extension == 'ts' || extension == 'js' || extension == 'map' || extension == 'php' || extension == 'json' || extension == 'xml'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/coding.png`;
  }
  if(extension == 'png' || extension == 'PNG' || extension == 'JPG' || extension == 'JPEG'  || extension == 'SVG' || extension == 'svg' || extension == 'jpg' || extension == 'jpeg' || extension == 'gif'){
    FileTypeImg = `${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/image.png`;
  }
  if(extension == "zip" || extension == "rar"){
    FileTypeImg=`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/zip.svg`;
  }


  return(
    <li className="clearfix">
      <div className="search-img-show">
        <img src={FileTypeImg} alt="image" />
      </div>
      <div className="search-contens-show">
        <a href={item.resource.webUrl} target='_blank' data-interception='off'> {DocName} </a>
        <h4> modified on {moment(item.resource.lastModifiedDateTime).format("DD MMM YYYY hh:mm a")}  </h4>              
        <p><Markup content={item.summary} /></p>
      </div>
    </li>
  );
}   
});



const EventsSearchResults: JSX.Element[] = this.state.EventsSearchResults.map(function (item, key) { 
  var handler=this; 
  $(".Loader").removeClass('open');    
  return(
    <li className="clearfix">
      <div className="search-img-show">
        <img src={`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/calendar.png`} alt="image" />
      </div>
      <div className="search-contens-show">
        <a href="#" style={{cursor:"default"}}> {item.resource.subject} </a>
        <h4>Start's at {moment(item.resource.start.dateTime).local().format("DD MMM YYYY hh:mm a")} - End's at {moment(item.resource.end.dateTime).local().format("DD MMM YYYY hh:mm a")}  </h4>
        <p> <Markup content={item.summary} /></p>
      </div>
    </li>
  );
});



const ExchangeMessageSearchResults: JSX.Element[] = this.state.MessageSearchResults.map(function (item, key) {
  var handler=this;
  $(".Loader").removeClass('open');
  if(item.resource.from != undefined && item.resource.from != null && item.resource.from != "undefined" && item.resource.from != "null"){
    if(item.resource.hasAttachments == true){
      let msgID = item.resource.conversationId;          
        return(
          <li className="clearfix">
            <div className="search-img-show">
              <img src={`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/outlook.png`} alt="image" />
            </div>
            <div className="search-contens-show">
              <a href={item.resource.webLink} target='_blank' data-interception='off'> {item.resource.subject} </a>
              <h4>From {item.resource.from.emailAddress.name} received on {moment(item.resource.receivedDateTime).local().format("DD MMM YYYY hh:mm a")}  </h4>
              <p> <Markup content={item.summary} /></p>
              <a href={item.resource.webLink} target='_blank' data-interception='off' className="attachemtscls"><i className="fa fa-paperclip" aria-hidden="true"></i>
              <span>view Attachement </span></a>
            </div>
          </li>
        );        
    }else{
      return(
        <li className="clearfix">
          <div className="search-img-show">
            <img src={`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/outlook.png`} alt="image" />
          </div>
          <div className="search-contens-show">
            <a href={item.resource.webLink} target='_blank' data-interception='off'> {item.resource.subject} </a>
            <h4>From {item.resource.from.emailAddress.name} received on {moment(item.resource.receivedDateTime).local().format("DD MMM YYYY hh:mm a")}  </h4>
            <p> <Markup content={item.summary} /></p>
          </div>
        </li>
      );
    }
  }else{
    if(item.resource.hasAttachments == true){                    
        return(
          <li className="clearfix">
            <div className="search-img-show">
              <img src={`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/outlook.png`} alt="image" />
            </div>
            <div className="search-contens-show">
              <a href={item.resource.webLink} target='_blank' data-interception='off'> {item.resource.subject} </a>
              <h4> received on {moment(item.resource.receivedDateTime).local().format("DD MMM YYYY hh:mm a")}  </h4>
              <p> <Markup content={item.summary} /></p>
            </div>
          </li>
        );        
    }else{
      return(
        <li className="clearfix">
          <div className="search-img-show">
            <img src={`${handler.props.absoluteURL}/SiteAssets/Search%20Assets/img/outlook.png`} alt="image" />
          </div>
          <div className="search-contens-show">
            <a href={item.resource.webLink} target='_blank' data-interception='off'> {item.resource.subject} </a>
            <h4>received on {moment(item.resource.receivedDateTime).local().format("DD MMM YYYY hh:mm a")}  </h4>
            <p> <Markup content={item.summary} /></p>
          </div>
        </li>
      );
    }
  }
});


    return(
      <div className={ styles.taqeefDefinitions }>
        <div className="container-block">
          <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/bannerimg.png" alt="MicrosoftTeams-image" className="top-banner"/>
          <div className="top-left top-left-logo">
            <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/logo.png" className="logo" alt="MicrosoftTeams-image"/>
          </div>

          <div className="centered">
            <h2 className="margin-0"><b>Definitions and Products</b></h2>
            <div className="topnav">
              <div className="inputborder">
                <div className="input-icons">
                 <i className="fa fa-search"></i>
                  <input className="input-field" id="txt-Search" placeholder="Search.." type="text" onChange={()=>this.iconRemove()} onKeyDown={(e)=>this.SearchResult(e)} autoComplete="off"/>
                  <i className="fa fa-close close-icon" onClick={()=>this.ClearSearchInput()}></i>  
                  <h6 className="err-msg"style={{display:"none",color:"red"}} id="txt-err-msg-search">Type something to search</h6>
                </div>
              </div>
              <div className="icon-bar">           
                <a href="#" className="def-class" data-tip data-for={"React-tooltip-definition"} data-custom-class="tooltip-custom" onClick={()=>this.DefinitionMasterSerach()}>
                  <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/definition.svg"  alt="definition-image"/>
                </a>
                <ReactTooltip id={"React-tooltip-definition"} place="bottom" type="dark" effect="solid">
                  <span>Definitions</span>
                </ReactTooltip>
                <a href="#" className="prod-class" data-tip data-for={"React-tooltip-product"} data-custom-class="tooltip-custom" onClick={()=>this.ProductSearchMaster()}>
                  <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/product%20(1).svg" alt="product-image"/>
                </a>
                <ReactTooltip id={"React-tooltip-product"} place="bottom" type="dark" effect="solid">
                  <span>Products</span>
                </ReactTooltip>   
                <a className="all-search" href="#" data-tip data-for={"React-tooltip-all"}  data-custom-class="tooltip-custom">
                  <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/menu.png" alt="MicrosoftTeams-image" onClick={()=>this.getEnteredQueryString()}/>
                </a>
                <ReactTooltip id={"React-tooltip-all"} place="bottom" type="dark" effect="solid">
                  <span>All</span>
                </ReactTooltip> 
                <a href="#" data-tip data-for={"React-tooltip-onedrive"} data-custom-class="tooltip-custom" className="anchor-drive" >
                  <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/cloud%20(1).png" alt="MicrosoftTeams-image"/>
                </a>
                <ReactTooltip id={"React-tooltip-onedrive"} place="bottom" type="dark" effect="solid">
                  <span>OneDrive</span>
                </ReactTooltip>
                <a href="#" data-tip data-for={"React-tooltip-sharepoint"} data-custom-class="tooltip-custom" className="anchor-sp" >
                  <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/share%20(1).png" alt="MicrosoftTeams-image"/>
                </a>
                <ReactTooltip id={"React-tooltip-sharepoint"} place="bottom" type="dark" effect="solid">
                  <span>SharePoint</span>
                </ReactTooltip>
                <a href="#" data-tip data-for={"React-tooltip-calendar"} data-custom-class="tooltip-custom" className="anchor-calendar" >
                  <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/calendar.png"className= "zoom" alt="MicrosoftTeams-image"/>
                </a>
                <ReactTooltip id={"React-tooltip-calendar"} place="bottom" type="dark" effect="solid">
                  <span>Calendar</span>
                </ReactTooltip>
                <a href="#" data-tip data-for={"React-tooltip-email"} data-custom-class="tooltip-custom" className="anchor-email" >
                  <img src="https://tmxin.sharepoint.com/sites/POC/taqeefIntranet/Definitions/SiteAssets/DefinitionAssets/Image/document.png" className= "zoom" alt="MicrosoftTeams-image"/>
                </a>
                <ReactTooltip id={"React-tooltip-email"} place="bottom" type="dark" effect="solid">
                  <span>EMail</span>
                </ReactTooltip>
              </div>
            </div>
          </div>
        </div>

        <div className="main-wrap">          
            <div className="result-block-wrap">
              <p className="margin-top">{this.state.value.length+this.state.Data.length+" Results Found"}</p>
              <ul> 
                  <li>
                  <select name="assetgroup" id="def-Department" className="def-Department form-control responsive-right-drpdwn"
                  onChange={()=>this.masterdefinitionfilter()}>
                  <option value=""disabled selected>--Department--</option>                   
                    {DefinitionDepartmentOptions}
                  </select></li>
                  <li>
                  <select name="assetgroup" id="def-Division" className="def-Division form-control"
                  onChange={()=>this.masterdefinitionfilter()}
                  >
                  <option value=""disabled selected>--Division--</option>                   
                    {DefinitionDivisionOptions}
                  </select></li>
                  {/*<li> 
                  <select name="assetgroup" id="def-Tag" className="def-Tag form-control responsive-right-drpdwn"
                  onChange={()=>this.masterdefinitionfilter()}>
                  <option value=""disabled selected>--Tag--</option>                  
                    {DefinitionTagOptions}
                  </select> </li>*/}
                 <li>
                  <select name="assetgroup" id="Product-Segment" className="Product-Segment form-control drpdwn-product"
                  onChange={()=>this.masterProductionfilter()}
                  >
                  <option value=""disabled selected>--Product--</option>                   
                    {ProductSegmentOptions}
                  </select> </li>
                  <li>
                  <select name="assetgroup" id="Product-Variant" className="Product-Variant form-control responsive-variant-drpdwn"
                  onChange={()=>this.masterProductionfilter()}
                  >
                  <option value=""disabled selected>--Variant--</option>                    
                    {ProductVariantOptions}
                  </select> </li>
                  <li>
                  <select name="assetgroup" id="Product-Category" className="Product-Category form-control" 
                  onChange={()=>this.masterProductionfilter()}
                  >                    
                   <option value=""disabled selected>--Category--</option>
                   {ProductCategoryOptions}
                  </select> </li>
                  <li className="reset-btn-wrap">
                  <div className="Filter_button">
                  <button className="btn-reset"  onClick={()=>this.filterClear()}><i className="fa fa-close clear-btn"></i>Reset</button>
                </div>
                  </li>
               
                
              </ul>
              <div className="Pagination-element-wrap">
                <p className="sub-head color-blue">{SearchResults}</p>              
              </div>            
              <div className="Pagination-element-wrap">
                {ProductResults}                                
              </div>   
              <div id="no-result" className="no-result-err" style={{display:"none"}}><i className="fa fa-warning"></i><h6>No result found!!!</h6></div>
              <div id="Search-err" className="no-result-err" style={{display:"none"}}><i className="fa fa-warning"></i><h6>Minimum 3 characters are required to search</h6></div>  
              <nav className="d-flex justify-content-center pagination-wrap">
                <ul id="numbers" className="pagination pagination-base pagination-boxed pagination-square mb-0">                          
                                                   
                </ul>
              </nav>                             
            </div>
          </div>
        </div>
    );
  }
}
function displayRows(index) {
  var start = (index - 1) * rowsPerPage;
  var end = start + rowsPerPage;
  
  // Hide all rows.
  rows.hide();
  
  // Show the proper rows for this page.
  rows.slice(start, end).show();
}


