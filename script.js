// Function to load and display data from the text file
function loadData() {
  // Make an XMLHttpRequest to load the text file
  var xhr = new XMLHttpRequest();
  xhr.open('GET', 'extracted_data.txt', true);

  xhr.onreadystatechange = function () {
    if (xhr.readyState === 4 && xhr.status === 200) {
      // Split the content of the text file into an array of lines
      var lines = xhr.responseText.split('\n');

      // Get the data-container element where we will display the data
      var dataContainer = document.getElementById('data-container');

      // Initialize a variable to keep track of the current card
      var currentCard = null;

      // Loop through the lines and display them on the webpage
      for (var i = 0; i < lines.length; i += 5) {
        var day = lines[i];
        var link = lines[i + 1];
        var location = lines[i + 2];
        var second_time = lines[i + 3];
        var time = lines[i + 4];

        // Create a card element for each day/group
        var card = document.createElement('div');
        card.className = 'card';

        // Create elements for day, time, and second time
        var dayElement = document.createElement('h2');
        dayElement.textContent = day;

        // Create an anchor element for the link
        var linkElement = document.createElement('a');
        linkElement.href = link;
        linkElement.target = '_blank'; // Open in a new tab
        linkElement.textContent = link;

        var locationElement = document.createElement('h3');
        locationElement.textContent = location;

        var secondTimeElement = document.createElement('p');
        secondTimeElement.textContent = second_time;

        var timeElement = document.createElement('p');
        timeElement.textContent = time;

        // Append elements to the card
        card.appendChild(dayElement);
        card.appendChild(linkElement);
        card.appendChild(locationElement);
        card.appendChild(secondTimeElement);
        card.appendChild(timeElement);

        // Append the card to the data container
        dataContainer.appendChild(card);
      }
    }
  };

  xhr.send();
}

// Call the loadData function when the page loads
window.addEventListener('load', loadData);
