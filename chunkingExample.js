(function(){"use strict"

	//this file will sum the multi-value choice field for every
	//	item in a list even if it exceeds the list view threshold.
	//	it assumes the library is "Documents" with a Multi-value
	//	choice field called "Document_Type"

	//it stores the result in the object "types"

	//a site or sub site under the top page a sharepoint
	var web="https://your.SharePoint.com/sites/yourweb/";
	//the name of the list, if it changes, this has to be updated
	var list="Documents";
	var activeQueriesAllowed=3;
	var chunkSize=3000;
	var dontCount=["Information Papers", "Ops Tasks"]

	var unstartedQueries=[],activeQueries=[],types={};

	function createChunkFilter(lowerBound,upperBound){
		return "$top="+chunkSize+"&$select=ID,Document_Type&$filter=ID le "+upperBound+" and ID gt "+lowerBound
	}

	function doWhenOneIsDone(queryResult){
		//make sure it's at least returns a valid array of results
		if(
			queryResult && "d" in queryResult &&
			"results" in queryResult.d &&
			Array.isArray(queryResult.d.results)
		){
			//loop through the results
			var results=queryResult.d.results;
			for(var i in results){
				var row=results[i];
				var docType=row.Document_Type;

				if(docType && "results" in docType){
					//get the year it was created
					var typesOnItem=docType.results;
					for(var i in typesOnItem){
						var oneOfTheTypes=typesOnItem[i];
						//don't count it if it's in the dontCount array
						if(dontCount.indexOf(oneOfTheTypes)==-1){
							if(oneOfTheTypes in types==false){
								types[oneOfTheTypes]=1
							}else{
								types[oneOfTheTypes]++
							}
						}
					}
				}
			}
		}
	}

	function doWhenAllAreDone(){
		$(document).ready(function(){
			var json=JSON.stringify(types, undefined, 4)
			//format the json so it is readable
			json=json.replace(/{/gi,"{<br />")
			json=json.replace(/}/gi,"<br />}")
			json=json.replace(/\,/gi,",<br />")
			//add the json to the page
			$("#DeltaPlaceHolderMain").html(json)
		})
	}

	//the bare bones basics for a rest call
	function rawRestRead(url){
		var doit={
			url: url,type: "Get",contentType: "application/json;odata=verbose",
			headers: {"Accept": "application/json; odata=verbose"},
		}
		return $.ajax(doit)
	}

	function getMostRecentIdForList(list){
		var dfd=$.Deferred();
		//do the rest call
		rawRestRead(web+"_api/web/lists/getbytitle('"+list+"')/Items?$select=ID&$top=1&$orderby=ID desc").then(function(queryResult){
			//make sure theree are results
			var results=queryResult.d.results
			//if there are, return the first records id ( should the the last item in the list ).  otherwise, return 0;
			if(results.length>0){
				var row=results[0];
				if(row && "ID" in row){
					dfd.resolve(row.ID)
				}else{dfd.resolve(0)}
			}else{dfd.resolve(0)}
		},function(){
			dfd.resolve(0)
		})
		return dfd.promise();
	}

	//exectute the chunking query
	function readItem(query){
		return rawRestRead(web+"_api/web/lists/getbytitle('"+list+"')/Items?"+query)
	}
	function getNextQueryForIndex(index){
		//consider it done until you are sure you have something to load there
		activeQueries[index]=null;

		//make sure some queries are left to do
		if(unstartedQueries.length>0){

			//grab a potential query
			var next=unstartedQueries.pop()

			//execute the query
			var newActiveQuery=readItem(next);

			//store a reference to the executed query
			activeQueries[index]=newActiveQuery;

			newActiveQuery.always(function(){

				//call the doWhenOneIsDone function as though the query called it
				var possiblePromise=doWhenOneIsDone.apply(this||{},arguments)

				//the on done function might return a promise, if so, wait for it to conclude before moving on to the next query.
				if(possiblePromise && "always" in possiblePromise){
					possiblePromise.always(function(){
						getNextQueryForIndex(index)
					})
				}else{
				//if it wasn't a promise, then go ahead and move on to the next query
					getNextQueryForIndex(index)
				}
			})
		}else{
			//they are all done untill you find one that isn't
			var allDone=true;
			for(var i in activeQueries){
				if(activeQueries[i]!==null){
					allDone=false;
				}
			}
			//they were all done, call the last function
			if(allDone){
				doWhenAllAreDone();
			}
		}

	}

	function start(){

		getMostRecentIdForList(list).then(function(id){
			//loop thorugh all of the items by chunksize
			for(var i=id;i>=0;i=i-chunkSize){
				var upperBound=i;
				var lowerBound=upperBound-chunkSize;
				//create the chunk filter
				var filter=createChunkFilter(lowerBound,upperBound)
				//add the chunk filter to the array
				unstartedQueries.push(filter);
			}

			//start a query in each open slot
			for(var i=0; i<activeQueriesAllowed;i++){
				getNextQueryForIndex(i)
			}
		})
	}
	start()
})()
