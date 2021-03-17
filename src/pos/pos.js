// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {
  ma, dma, ema, sma, wma
} from 'moving-averages';

const d3 = require("d3");

var options = {"readability": true, "words": true, "average": true, "z-scores": false, "detail": 100};

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("refresh").onclick = refresh;

    refresh(true);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});


function refresh(first = false) {
  Word.run(function (context) {
      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");

      return context.sync()
        .then(function() {
          draw_chart(paragraphs);
        })
        .then(context.sync);
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function draw_chart(paragraphs) {
  d3.select("svg").remove();

  // set the dimensions and margins of the graph
  var margin = {top: 30, right: 30, bottom: 50, left: 90}
  , width = window.innerWidth - margin.left - margin.right // Use the window's width 
  , height = (5 * 50) - margin.top - margin.bottom; // Use the window's height

  var colourScale = d3.scaleSequential()
  .domain([0,21])
  .interpolator(d3.interpolateYlGnBu);

  // append the svg object to the body of the page
  var svg = d3.select("#pos_vis")
  .append("svg")
  .attr("width", width + margin.left + margin.right)
  .attr("height", height + margin.top + margin.bottom)
  .append("g")
  .attr("transform",
        "translate(" + margin.left + "," + margin.top + ")");

  var data = [{"name": "Adverbs", "count": 1},
  {"name": "Adjectives", "count": 6},
  {"name": "Verbs", "count": 14},
  {"name": "Pronouns", "count": 21},
  {"name": "Proper Nouns", "count": 9}];


  // Add X axis
  var x = d3.scaleLinear()
  .domain([0, 21])
  .range([ 0, width]);

  // Y axis
  var y = d3.scaleBand()
  .range([ 0, height ])
  .domain(data.map(function(d) { return d.name; }))
  .padding(.1);

  //Bars
  svg.selectAll("myRect")
  .data(data)
  .enter()
  .append("rect")
  .attr("x", x(0) )
  .attr("y", function(d) { return y(d.name); })
  .attr("width", function(d) { return x(d.count); })
  .attr("height", y.bandwidth() )
  .attr("fill", d => colourScale(d.count));

  svg.append("g")
  .attr("transform", "translate(0," + height + ")")
  .call(d3.axisBottom(x))
  .selectAll("text")
    .attr("transform", "translate(-10,0)rotate(-45)")
    .style("text-anchor", "end");

  svg.append("g")
  .call(d3.axisLeft(y));

      // text label for the x axis
      svg.append("text")             
      .attr("transform",
            "translate(" + (width/2) + " ," + 
                           (height + 40) + ")")
      .style("text-anchor", "middle")
      .text("Count");


  // .attr("x", function(d) { return x(d.Country); })
  // .attr("y", function(d) { return y(d.Value); })
  // .attr("width", x.bandwidth())
  // .attr("height", function(d) { return height - y(d.Value); })
  // .attr("fill", "#69b3a2")
}