// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

const d3 = require("d3");

import {
  ma, dma, ema, sma, wma
} from 'moving-averages';

var messageQueue = [];

const worker = new Worker("pov_worker.js");

var pronoun_data = {"1st": 0, "2nd": 0, "3rd": 0};

var entities = [];

var lineChartData = [];

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";

  updatePronounData(e.data.pronouns);
  lineChartData.push(e.data.pronouns);

  entities = entities.concat(e.data.entities);

  redrawPieChart(messageQueue.length == 0);

  if(messageQueue.length > 0) {  
    worker.postMessage({"text": messageQueue.pop()});
  } else {
    let dictionary = mode(entities);

    addInfo(Object.entries(pronoun_data).sort((a,b) => a[1] - b[1])[2][0], dictionary);

    lineChartData.reverse();
    drawLineChart("1st");
  }
};

function mode(array)
{
    let entities = {};

    array.forEach((entity) => {
      if(entity in entities) {
        entities[entity] += 1;
      } else {
        entities[entity] = 1;
      }
    });

    return Object.entries(entities).sort((a,b) => b[1] - a[1]);
}

var margin = {top: 30, right: 30, bottom: 30, left: 30}
, width = window.innerWidth - margin.left - margin.right
, height = window.innerWidth - margin.top - margin.bottom;

var radius = Math.min(width, height) / 2 - margin.top;

var svg = d3.select("#pie_chart_vis")
  .append("svg")
  .attr("width", width + margin.left + margin.right)
  .attr("height", height + margin.top + margin.bottom)
  .append("g")
  .attr("transform",
        "translate(" + ((width / 2) + margin.left) + "," + ((height / 2) + margin.top) + ")");

var colourScale = d3.scaleOrdinal(d3.schemeSet3);

colourScale.domain(Object.keys(pronoun_data));

var pie = d3.pie()
  .value((d) => d.value);

var arcGenerator = d3.arc()
  .innerRadius(0)
  .outerRadius(radius);

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.

    document.getElementById("refresh").onclick = refresh;
    document.getElementById("line_key").onchange = onLineKeyChange;

    refresh();
    drawPieChart();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function onLineKeyChange() {
  redrawLineChart(document.getElementById("line_key").value);
}

function loading() {
  document.getElementById("summary").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("entities").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("line_chart_vis").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

function addInfo(person, dictionary) {
  document.getElementById("summary").innerHTML = `
  <div class="ms-MessageBar ms-MessageBar--success">
    <div class="ms-MessageBar-content">
      <div class="ms-MessageBar-icon">
        <i class="ms-Icon ms-Icon--Completed"></i>
      </div>
      <div class="ms-MessageBar-text">
        The selected text is most likely in <span style="font-weight: 700">` + person + ` person</span>, with <span style="font-weight: 700">` + dictionary[0][0] + 
        `</span> as the perspective character.
      </div>
    </div>
  </div>`;

  document.getElementById("entities").innerHTML = `
  <div class="ms-MessageBar">
    <div class="ms-MessageBar-content">
      <div class="ms-MessageBar-icon">
        <i class="ms-Icon ms-Icon--Info"></i>
      </div>
      <div class="ms-MessageBar-text">
        The most common entities in this text are <span style="font-weight: 700">` + dictionary[0][0] + `</span>, <span style="font-weight: 700">`
         + dictionary[1][0] + `</span> and <span style="font-weight: 700">` + dictionary[2][0] +`</span>.
      </div>
    </div>
  </div>`;
}

function refresh() {
  loading();
  resetPronounData();

  Word.run(function (context) {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");

    var selection = context.document.getSelection();

    selection.load("text");

    selection.paragraphs.load("text");

    return context.sync()
      .then(function() {
        document.getElementById("debug").innerHTML = "Message sending...";
        
        if(selection.text.length == 0) {
          messageQueue = paragraphs.items.map(paragraph => paragraph.text);
        } else {
          let results = selection.paragraphs.items.map(paragraph => paragraph.text);

          if(results.length > 1) {
            let wholeText = results.join('\r');
            let match = new RegExp('(.*)' + selection.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(wholeText);

            if(match != null) {
              let firstParagraphMatch = new RegExp(match[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(results[0]);
              results[0] = firstParagraphMatch[1];

              let lastParagraphMatch = new RegExp('(.*)' + match[2].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g').exec(results[results.length-1]);
              results[results-1] = lastParagraphMatch[1];
            }
          } else {
            results[0] = selection.text;
          }

          messageQueue = results;
        }
        
        worker.postMessage({"text": messageQueue.pop()});
        document.getElementById("debug").innerHTML = "Message sent.";
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

function drawPieChart() {
  var data_ready = pie(Object.entries(pronoun_data).map(function(x) {return {"key": x[0], "value": x[1]}; }));

  svg
  .selectAll('mySlices')
  .data(data_ready)
  .enter()
  .append('path')
  .attr("class", "arc")
  .attr('d', arcGenerator)
  .attr('fill', (d) => colourScale(d.data.key))
  .attr("stroke", "white")
  .style("stroke-width", "2px");

svg
  .selectAll('mySlices')
  .data(data_ready)
  .enter()
  .append('text')
  .attr("class", "label")
  .text(function(d){ return d.data.key})
  .attr("transform", function(d) { return "translate(" + arcGenerator.centroid(d) + ")";  })
  .style("text-anchor", "middle")
  .style("font-size", 17)
  .style("opacity", d => d.data.value > 0 ? 100 : 0);
}

function redrawPieChart(animate) {
  var data_ready = pie(Object.entries(pronoun_data).map(function(x) {return {"key": x[0], "value": x[1]}; }));

  svg
    .selectAll('.arc')
    .data(data_ready)
    .transition().duration(animate ? 800 : 1)
    .attr('d', arcGenerator)
    .attr('fill', (d) => colourScale(d.data.key))
    .attr("stroke", "white")
    .style("stroke-width", "2px");

  svg
    .selectAll('.label')
    .data(data_ready)
    .transition().duration(animate ? 800 : 10)
    .text(function(d){ return d.data.key})
    .attr("transform", function(d) { return "translate(" + arcGenerator.centroid(d) + ")";  })
    .style("text-anchor", "middle")
    .style("font-size", 17)
    .style("opacity", d => d.data.value > 0 ? 100 : 0);
}

function updatePronounData(newData) {
  pronoun_data["1st"] += newData["1st"];
  pronoun_data["2nd"] += newData["2nd"];
  pronoun_data["3rd"] += newData["3rd"];
}

function resetPronounData() {
  pronoun_data = {"1st": 0, "2nd": 0, "3rd": 0};
  lineChartData = [];
}

function movingAverage(data, width) {
  return Array.from(ma(data, width));
}

function drawLineChart(key) {
  document.getElementById("line_chart_vis").innerHTML = "";

  var line_margin = {top: 50, right: 50, bottom: 50, left: 60}
    , line_width = window.innerWidth - line_margin.left - line_margin.right
    , line_height = window.innerHeight - line_margin.top - line_margin.bottom;

  var wordCounts = lineChartData.map(d => d["1st"] + d["2nd"] + d["3rd"]);

  var densities = lineChartData.map((d, i) => wordCounts[i] > 0 ? (d[key] / wordCounts[i]) : 0);

  let averages = movingAverage(densities, Math.ceil(densities.length * 0.15));

  var xScaleLine = d3.scaleLinear()
    .domain([0, d3.max(densities)])
    .range([0, line_width]);

  var yScaleLine = d3.scaleLinear()
      .domain([lineChartData.length-1,0])
      .range([line_height, 0]);

  var line = d3.line()
    .x(function(d) { return xScaleLine(d); })
    .y(function(d, i) { return yScaleLine(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d));

  d3.select("#line_chart_vis").select("svg").remove();

  var line_svg = d3.select("#line_chart_vis").append("svg")
    .attr("width", line_width + line_margin.left + line_margin.right)
    .attr("height", line_height + line_margin.top + line_margin.bottom)
    .append("g")
      .attr("transform", "translate(" + line_margin.left + "," + line_margin.top + ")");

  line_svg.append("g")
      .attr("class", "y-axis-line")
      .call(d3.axisLeft(yScaleLine));

  line_svg.append("path")
      .datum(densities)
      .attr("class", "line")
      .attr("stroke", "#0084ff")
      .attr("fill", "none")
      .attr("d", line);

  line_svg.append("path")
    .datum(averages)
    .attr("class", "lineAverage")
    .attr("d", line);

  line_svg.append("g")
    .attr("class", "x-axis-line")
    .attr("transform", "translate(0,0)")
    .call(d3.axisTop(xScaleLine).ticks(5));

  line_svg.append("text")             
  .attr("transform",
        "translate(" + (line_width/2) + " ," + 
                       (- 30) + ")")
  .style("text-anchor", "middle")
  .text("Density");

  line_svg.append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 0 - line_margin.left + 10)
      .attr("x",0 - (line_height / 2))
      .attr("dy", "1em")
      .style("text-anchor", "middle")
      .text("Paragraph #");
}

function redrawLineChart(key) {
  var wordCounts = lineChartData.map(d => d["1st"] + d["2nd"] + d["3rd"]);

  var densities = lineChartData.map((d, i) => wordCounts[i] > 0 ? (d[key] / wordCounts[i]) : 0);

  let averages = movingAverage(densities, Math.ceil(densities.length * 0.15));

  var svg = d3.select("#line_chart_vis").select("svg");
  
  var margin = {top: 50, right: 50, bottom: 50, left: 60}
    , width = window.innerWidth - margin.left - margin.right
    , height = window.innerHeight - margin.top - margin.bottom;

  var xScaleLine = d3.scaleLinear()
    .domain([0, d3.max(densities)])
    .range([0, width]);

  var yScaleLine = d3.scaleLinear()
    .domain([lineChartData.length-1,0])
    .range([height, 0]);

  var line = d3.line()
    .x(function(d) { return xScaleLine(d); })
    .y(function(d, i) { return yScaleLine(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d));

  // redraw the line
  d3.selectAll(".line")
      .datum(densities)
      .transition().duration(800)
      .attr("d", line);

  d3.selectAll(".lineAverage")
    .datum(averages)
    .transition().duration(800)
    .attr("d", line);
  
  // redraw the x axis with the new scale
  svg.selectAll(".x-axis-line").remove();

  svg.append("g")
    .attr("class", "x-axis-line")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")")
    .call(d3.axisTop(xScaleLine).ticks(5));
}