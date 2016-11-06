var officegen = require('officegen');
var async = require('async');
var fs = require('fs');
var express = require('express');
var app = express();
var mongoose = require('mongoose');
var bodyParser = require('body-parser');

app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({extended: true}));
app.use(express.static(__dirname + '/public'));


var PORT = process.env.PORT || '3000';

mongoose.connect('mongodb://localhost/wordgen1');

var PersonSchema = new mongoose.Schema({
	fullname: String,
	age: Number,
	address: String,
	course: String
});

var Person = mongoose.model('Person', PersonSchema);

app.get('/', function(req, res){
	res.render('homepage');
});

app.post('/addperson', function(req, res){
	var newPerson = req.body.person;

	Person.create(newPerson, function(err, thePerson){
		if(err){
			console.log(err);
		} else {
			res.redirect('back');
		}
	});
});


app.post('/genword', function(req, res){
	Person.findOne({age: 25}, function(err, thePerson){
		var docx = officegen('docx');

		var pObj = docx.createP({align: 'right'});
		pObj.addText(thePerson.fullname.toString(), {color: 'red'});

		var theText = docx.createP({align:'left'});
		theText.addText(thePerson.age.toString());

		var out = fs.createWriteStream ( './public/john.docx' );

		out.on ( 'error', function ( err ) {
			console.log ( err );
		});

		async.parallel ([
			function ( done ) {
				out.on ( 'close', function () {
					console.log ( 'Created a CV for ' + thePerson.fullname );
					done ( null );
				});
				docx.generate ( out );
			}

		], function ( err ) {
			if ( err ) {
				console.log ( 'error: ' + err );
			} 
		});
		res.redirect('back');
	});
});




app.post('/instantgenword', function(req, res){
	var newPerson = req.body.person;

	Person.create(newPerson, function(err, thePerson){
		if(err){
			console.log(err);
		} else {
			var docx = officegen('docx');

			var pObj = docx.createP();
			pObj.addText(thePerson.fullname.toString(), {bold: true});

			var theText = docx.createP({align:'left'});
			theText.addText(thePerson.age.toString());
			theText.addLineBreak();
			theText.addText(thePerson.address.toString());
			theText.addLineBreak();
			theText.addText(thePerson.course.toString());

			var out = fs.createWriteStream ( './public/john.docx' );

			out.on ( 'error', function ( err ) {
				console.log ( err );
			});

			async.parallel ([
				function ( done ) {
					out.on ( 'close', function () {
						console.log ( 'Created a CV for ' + thePerson.fullname );
						done ( null );
					});
					docx.generate ( out );
				}

			], function ( err ) {
				if ( err ) {
					console.log ( 'error: ' + err );
				} 
			});

			res.redirect('back');
		}
	});
});











//Word generator
var docx = officegen('docx');
var pObj = docx.createP({align: 'right'});

pObj.addText('Hey there');


var theText = docx.createP({align:'left'});

theText.addText("This is my name!");



// var out = fs.createWriteStream ( 'john.docx' );

// out.on ( 'error', function ( err ) {
// 	console.log ( err );
// });

// async.parallel ([
// 	function ( done ) {
// 		out.on ( 'close', function () {
// 			console.log ( 'Finish to create a DOCX file.' );
// 			done ( null );
// 		});
// 		docx.generate ( out );
// 	}

// ], function ( err ) {
// 	if ( err ) {
// 		console.log ( 'error: ' + err );
// 	} 
// });


app.listen(PORT, function(req, res){
	console.log('Server started at Port ' + PORT);
});