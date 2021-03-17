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

    refresh();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function refresh() {
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

  var margin = {top: 50, right: 50, bottom: 50, left: 50}
  , width = window.innerWidth - margin.left - margin.right // Use the window's width 
  , height = window.innerWidth - margin.top - margin.bottom; // Use the window's height

    // append the svg object to the body of the page
  var svg = d3.select("#sentiment_vis")
    .append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
    .attr("transform",
          "translate(" + margin.left + "," + margin.top + ")");
  
    let radialScale = d3.scaleLinear()
      .domain([0,10])
      .range([0,width/2]);

    let ticks = [2,4,6,8,10];

    ticks.forEach(t =>
      svg.append("circle")
      .attr("cx", width/2)
      .attr("cy", width/2)
      .attr("fill", "none")
      .attr("stroke", "gray")
      .attr("r", radialScale(t))
    );

    // ticks.forEach(t =>
    //   svg.append("text")
    //   .attr("x", width/2 + 5)
    //   .attr("y", width/2 - radialScale(t))
    //   .text(t.toString())
    // );

    let features = ["Joy", "Sadness", "Anger", "Fear", "Disgust"];
    let data = [{"Joy": 2.0, "Sadness": 9.0, "Anger": 7.0, "Fear": 5.0, "Disgust": 2.5}];

    function angleToCoordinate(angle, value){
      let x = Math.cos(angle) * radialScale(value);
      let y = Math.sin(angle) * radialScale(value);
      return {"x": width/2 + x, "y": width/2 - y};
    }

    for (var i = 0; i < features.length; i++) {
      let ft_name = features[i];
      let angle = (Math.PI / 2) + (2 * Math.PI * i / features.length);
      let line_coordinate = angleToCoordinate(angle, 10);
      let label_coordinate = angleToCoordinate(angle, 10.5);
  
      //draw axis line
      svg.append("line")
      .attr("x1", width/2)
      .attr("y1", width/2)
      .attr("x2", line_coordinate.x)
      .attr("y2", line_coordinate.y)
      .attr("stroke","black");
  
      //draw axis label
      svg.append("text")
      .attr("x", label_coordinate.x)
      .attr("y", label_coordinate.y)
      .text(ft_name);

      let line = d3.line()
        .x(d => d.x)
        .y(d => d.y);

      let colors = ["darkorange", "gray", "navy"];

      function getPathCoordinates(data){
        let coordinates = [];
        for (var i = 0; i < features.length; i++){
            let ft_name = features[i];
            let angle = (Math.PI / 2) + (2 * Math.PI * i / features.length);
            coordinates.push(angleToCoordinate(angle, data[ft_name]));
        }

        coordinates.push(coordinates[0])
        return coordinates;
      }

      let coordinates = getPathCoordinates(data[0]);

      svg.append("path")
        .datum(coordinates)
        .attr("d",line)
        .attr("stroke-width", 3)
        .attr("stroke", "#0084ff")
        .attr("fill", "#accdeb")
        .attr("stroke-opacity", 1)
        .attr("opacity", 0.5);
  }
}