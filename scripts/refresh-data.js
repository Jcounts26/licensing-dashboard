const https = require('https');
const fs = require('fs');
const path = require('path');

// Configuration
const ORG = process.env.AZURE_DEVOPS_ORG || 'mi-devops';
const PROJECT = process.env.AZURE_DEVOPS_PROJECT || 'Mi-Case_eLicensing';
const PAT = process.env.AZURE_DEVOPS_PAT;
const SPRINT_INPUT = process.env.SPRINT_NAME || 'current';

if (!PAT) {
  console.error('Error: AZURE_DEVOPS_PAT environment variable is required');
  process.exit(1);
}

const AUTH_HEADER = 'Basic ' + Buffer.from(':' + PAT).toString('base64');

// Helper to make Azure DevOps API requests
function apiRequest(urlPath) {
  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'dev.azure.com',
      path: urlPath,
      method: 'GET',
      headers: {
        'Authorization': AUTH_HEADER,
        'Content-Type': 'application/json'
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(JSON.parse(data));
        } else {
          reject(new Error(`API request failed: ${res.statusCode} - ${data}`));
        }
      });
    });

    req.on('error', reject);
    req.end();
  });
}

// Get team iterations to find current sprint
async function getIterations() {
  const url = `/${ORG}/${PROJECT}/Licensing/_apis/work/teamsettings/iterations?api-version=7.0`;
  const result = await apiRequest(url);
  return result.value || [];
}

// Get work items for an iteration
async function getIterationWorkItems(iterationPath) {
  // WIQL query to get work items in the iteration assigned to developers
  const wiql = {
    query: `SELECT [System.Id], [System.Title], [System.State], [System.AssignedTo], [Microsoft.VSTS.Scheduling.StoryPoints], [System.WorkItemType]
            FROM WorkItems
            WHERE [System.IterationPath] = '${iterationPath}'
            AND [System.WorkItemType] IN ('User Story', 'Bug')
            ORDER BY [System.AssignedTo]`
  };

  const url = `/${ORG}/${PROJECT}/_apis/wit/wiql?api-version=7.0`;

  return new Promise((resolve, reject) => {
    const options = {
      hostname: 'dev.azure.com',
      path: url,
      method: 'POST',
      headers: {
        'Authorization': AUTH_HEADER,
        'Content-Type': 'application/json'
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode >= 200 && res.statusCode < 300) {
          resolve(JSON.parse(data));
        } else {
          reject(new Error(`WIQL query failed: ${res.statusCode} - ${data}`));
        }
      });
    });

    req.on('error', reject);
    req.write(JSON.stringify(wiql));
    req.end();
  });
}

// Get work item details in batches
async function getWorkItemDetails(ids) {
  if (ids.length === 0) return [];

  const batchSize = 200;
  const allItems = [];

  for (let i = 0; i < ids.length; i += batchSize) {
    const batch = ids.slice(i, i + batchSize);
    const idsParam = batch.join(',');
    const url = `/${ORG}/${PROJECT}/_apis/wit/workitems?ids=${idsParam}&fields=System.Id,System.Title,System.State,System.AssignedTo,Microsoft.VSTS.Scheduling.StoryPoints,System.WorkItemType&api-version=7.0`;
    const result = await apiRequest(url);
    allItems.push(...(result.value || []));
  }

  return allItems;
}

// Calculate progress based on work item states
function calculateProgress(items) {
  if (items.length === 0) return 0;

  const closedStates = ['Closed', 'Done', 'Resolved'];
  const inProgressStates = ['Ready for QA', 'In QA', 'Ready for Review'];

  let totalPoints = 0;
  let completedPoints = 0;

  items.forEach(item => {
    const points = item.fields['Microsoft.VSTS.Scheduling.StoryPoints'] || 0;
    const state = item.fields['System.State'];

    totalPoints += points;

    if (closedStates.includes(state)) {
      completedPoints += points;
    } else if (inProgressStates.includes(state)) {
      completedPoints += points * 0.75; // 75% complete if in QA
    }
  });

  return totalPoints > 0 ? Math.round((completedPoints / totalPoints) * 100) : 0;
}

// Group work items by developer
function groupByDeveloper(items) {
  const developers = {};

  items.forEach(item => {
    const assignedTo = item.fields['System.AssignedTo'];
    if (!assignedTo) return;

    const devName = assignedTo.displayName;

    if (!developers[devName]) {
      developers[devName] = {
        name: devName,
        items: [],
        points: 0,
        stories: 0,
        focusAreas: new Set()
      };
    }

    const points = item.fields['Microsoft.VSTS.Scheduling.StoryPoints'] || 0;
    const title = item.fields['System.Title'];

    developers[devName].items.push({
      id: item.id,
      title: title,
      points: points,
      state: item.fields['System.State'],
      type: item.fields['System.WorkItemType']
    });

    developers[devName].points += points;
    developers[devName].stories += 1;

    // Extract focus area from title (first few words or component name)
    const focusMatch = title.match(/^([^:-]+)/);
    if (focusMatch) {
      const focus = focusMatch[1].trim().substring(0, 20);
      if (focus.length > 3) {
        developers[devName].focusAreas.add(focus);
      }
    }
  });

  return developers;
}

// Update the HTML file with new data
function updateHtmlFile(sprintName, sprintPoints, devStories) {
  const htmlPath = path.join(__dirname, '..', 'index.html');
  let html = fs.readFileSync(htmlPath, 'utf8');

  // Find and update the sprintData object for the specific sprint
  const sprintDataRegex = new RegExp(
    `'${sprintName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}':\\s*\\{[^}]*sprintPoints:\\s*\\[[^\\]]*\\][^}]*devStories:\\s*\\{[^}]*(?:\\{[^}]*\\}[^}]*)*\\}[^}]*\\}`,
    's'
  );

  // Build the new sprint data
  const sprintPointsStr = JSON.stringify(sprintPoints, null, 20).replace(/"/g, "'").replace(/\n\s{20}/g, '\n                    ');
  const devStoriesStr = JSON.stringify(devStories, null, 24).replace(/"/g, "'").replace(/\n\s{24}/g, '\n                        ');

  const newSprintData = `'${sprintName}': {
                sprintPoints: ${sprintPointsStr},
                devStories: ${devStoriesStr}
            }`;

  // Try to replace existing sprint data
  if (sprintDataRegex.test(html)) {
    html = html.replace(sprintDataRegex, newSprintData);
  } else {
    console.log(`Sprint "${sprintName}" not found in HTML, may need manual update`);
  }

  // Update the refresh timestamp
  const timestampRegex = /Last data refresh: <span id="refreshTimestamp">[^<]*<\/span>/;
  const now = new Date().toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'long',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  });
  html = html.replace(timestampRegex, `Last data refresh: <span id="refreshTimestamp">${now}</span>`);

  fs.writeFileSync(htmlPath, html);
  console.log(`Updated HTML file with data for ${sprintName}`);
}

// Main function
async function main() {
  try {
    console.log('Fetching iterations from Azure DevOps...');
    const iterations = await getIterations();

    // Find the target sprint
    let targetIteration;
    const now = new Date();

    if (SPRINT_INPUT === 'current') {
      // Find current sprint based on date
      targetIteration = iterations.find(iter => {
        if (!iter.attributes) return false;
        const start = new Date(iter.attributes.startDate);
        const end = new Date(iter.attributes.finishDate);
        return now >= start && now <= end;
      });

      if (!targetIteration) {
        // If no current sprint, find the next upcoming one
        targetIteration = iterations
          .filter(iter => iter.attributes && new Date(iter.attributes.startDate) > now)
          .sort((a, b) => new Date(a.attributes.startDate) - new Date(b.attributes.startDate))[0];
      }
    } else {
      // Find specific sprint by name
      targetIteration = iterations.find(iter => iter.name === SPRINT_INPUT || iter.path.includes(SPRINT_INPUT));
    }

    if (!targetIteration) {
      console.error('Could not find target sprint');
      process.exit(1);
    }

    console.log(`Refreshing data for: ${targetIteration.name}`);
    console.log(`Path: ${targetIteration.path}`);

    // Get work items for the iteration
    console.log('Fetching work items...');
    const wiqlResult = await getIterationWorkItems(targetIteration.path);
    const workItemIds = (wiqlResult.workItems || []).map(wi => wi.id);

    console.log(`Found ${workItemIds.length} work items`);

    if (workItemIds.length === 0) {
      console.log('No work items found for this sprint');
      return;
    }

    // Get work item details
    console.log('Fetching work item details...');
    const workItems = await getWorkItemDetails(workItemIds);

    // Filter to only developers (exclude QA, PMs, etc.)
    const devNames = ['Dan Morris', 'Aparna Gupta', 'Vinay Patel', 'Sandip Pandya', 'Rajini Matharasi', 'Nosa Odaro', 'Nosa'];
    const devWorkItems = workItems.filter(item => {
      const assignedTo = item.fields['System.AssignedTo'];
      return assignedTo && devNames.some(name => assignedTo.displayName.includes(name));
    });

    console.log(`Found ${devWorkItems.length} developer work items`);

    // Group by developer
    const developers = groupByDeveloper(devWorkItems);

    // Build sprintPoints array
    const sprintPoints = Object.values(developers)
      .map(dev => ({
        name: dev.name,
        points: dev.points,
        stories: dev.stories,
        focus: Array.from(dev.focusAreas).slice(0, 3).join(', '),
        progress: calculateProgress(dev.items.map(i => ({ fields: { 'System.State': i.state, 'Microsoft.VSTS.Scheduling.StoryPoints': i.points } })))
      }))
      .sort((a, b) => b.points - a.points);

    // Build devStories object
    const devStories = {};
    Object.values(developers).forEach(dev => {
      devStories[dev.name] = dev.items;
    });

    console.log('\nSprint Summary:');
    console.log('===============');
    sprintPoints.forEach(dev => {
      console.log(`${dev.name}: ${dev.points} pts, ${dev.stories} stories, ${dev.progress}% complete`);
    });

    // Update HTML file
    updateHtmlFile(targetIteration.name, sprintPoints, devStories);

    console.log('\nRefresh complete!');

  } catch (error) {
    console.error('Error:', error.message);
    process.exit(1);
  }
}

main();
