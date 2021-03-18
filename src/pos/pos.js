// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

const d3 = require("d3");

const worker = new Worker("pos_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";
  draw_chart(e.data.pos);
  let sentences = e.data.active + e.data.passive;
  let active = ((e.data.active / sentences) * 100).toFixed(2);
  let passive = ((e.data.passive / sentences) * 100).toFixed(2);
  document.getElementById("active-passive").innerHTML = active > passive ? `<div class="ms-MessageBar ms-MessageBar--success">
  <div class="ms-MessageBar-content">
    <div class="ms-MessageBar-icon">
      <i class="ms-Icon ms-Icon--Completed"></i>
    </div>
    <div class="ms-MessageBar-text">
      <span style="font-weight: 700">` + active + `% Active</span> | ` + passive + `% Passive
    </div>
  </div>
</div>`
  : `<div class="ms-MessageBar ms-MessageBar--error">
  <div class="ms-MessageBar-content">
    <div class="ms-MessageBar-icon">
      <i class="ms-Icon ms-Icon--ErrorBadge"></i>
    </div>
    <div class="ms-MessageBar-text">
      ` + active + `% Active | <span style="font-weight: 700">` + passive + `% Passive</span>
    </div>
  </div>
</div>`;

  let words = e.data.pos.map(x => x.count).reduce((a,b) => a + b, 0);
  let adjectives = e.data.pos[0].count / words;
  let adverbs = e.data.pos[1].count / words;
  let pronouns = e.data.pos[4].count / words;

  document.getElementById("notifications").innerHTML = "";

  document.getElementById("notifications").innerHTML += adjectives <= 0.1 ? `<div class="ms-MessageBar ms-MessageBar--success" style="width: 100%">
    <div class="ms-MessageBar-content">
      <div class="ms-MessageBar-icon">
        <i class="ms-Icon ms-Icon--Completed"></i>
      </div>
      <div class="ms-MessageBar-text">
        That's a nice ratio of <span style="font-weight: 700">adjectives</span> you've got there!
      </div>
    </div>
  </div>` : `<div class="ms-MessageBar ms-MessageBar--warning" style="width: 100%">
    <div class="ms-MessageBar-content">
      <div class="ms-MessageBar-icon">
        <i class="ms-Icon ms-Icon--Info"></i>
      </div>
      <div class="ms-MessageBar-text">
        Woah! That's a lot of <span style="font-weight: 700">adjectives</span> there!
      </div>
    </div>
  </div>`;

  document.getElementById("notifications").innerHTML += adverbs <= 0.05 ? `<div class="ms-MessageBar ms-MessageBar--success" style="width: 100%">
  <div class="ms-MessageBar-content">
    <div class="ms-MessageBar-icon">
      <i class="ms-Icon ms-Icon--Completed"></i>
    </div>
    <div class="ms-MessageBar-text">
      That's a nice ratio of <span style="font-weight: 700">adverbs</span> you've got there!
    </div>
  </div>
</div>` : `<div class="ms-MessageBar ms-MessageBar--warning" style="width: 100%">
  <div class="ms-MessageBar-content">
    <div class="ms-MessageBar-icon">
      <i class="ms-Icon ms-Icon--Info"></i>
    </div>
    <div class="ms-MessageBar-text">
      Woah! That's a lot of <span style="font-weight: 700">adverbs</span> there!
    </div>
  </div>
</div>`;

document.getElementById("notifications").innerHTML += pronouns <= 0.2 ? `<div class="ms-MessageBar ms-MessageBar--success" style="width: 100%">
<div class="ms-MessageBar-content">
  <div class="ms-MessageBar-icon">
    <i class="ms-Icon ms-Icon--Completed"></i>
  </div>
  <div class="ms-MessageBar-text">
    That's a nice ratio of <span style="font-weight: 700">pronouns</span> you've got there!
  </div>
</div>
</div>` : `<div class="ms-MessageBar ms-MessageBar--warning" style="width: 100%">
<div class="ms-MessageBar-content">
  <div class="ms-MessageBar-icon">
    <i class="ms-Icon ms-Icon--Info"></i>
  </div>
  <div class="ms-MessageBar-text">
    Woah! That's a lot of <span style="font-weight: 700">pronouns</span> there!
  </div>
</div>
</div>`;
};

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

function loading() {
  document.getElementById("active-passive").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("pos_vis").innerHTML = `<br><div class="ms-Spinner"></div><br>`;
  document.getElementById("notifications").innerHTML = `<br><div class="ms-Spinner"></div><br>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

function refresh() {
  loading();

  Word.run(function (context) {
    let body = context.document.body;
    body.load("text");

    return context.sync()
      .then(function() {
        document.getElementById("debug").innerHTML = "Message sending...";
        worker.postMessage({"text": body.text});
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

function draw_chart(data) {
  document.getElementById("pos_vis").innerHTML = "";

  // set the dimensions and margins of the graph
  var margin = {top: 30, right: 30, bottom: 50, left: 90}
  , width = window.innerWidth - margin.left - margin.right // Use the window's width 
  , height = (5 * 50) - margin.top - margin.bottom; // Use the window's height

  var colourScale = d3.scaleSequential()
  .domain([0,d3.max(data.map(d => d.count))])
  .interpolator(d3.interpolateYlGnBu);

  // append the svg object to the body of the page
  var svg = d3.select("#pos_vis")
  .append("svg")
  .attr("width", width + margin.left + margin.right)
  .attr("height", height + margin.top + margin.bottom)
  .append("g")
  .attr("transform",
        "translate(" + margin.left + "," + margin.top + ")");


  // Add X axis
  var x = d3.scaleLinear()
  .domain([0, d3.max(data.map(d => d.count))])
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