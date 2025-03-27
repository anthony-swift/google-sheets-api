// importWorkouts.js

const axios = require('axios');

// Your Hevy API key
const API_KEY = 'bc112740-430b-4b00-beaa-9253a1f56c14';

/**
 * Fetches workouts from Hevy for a given page with the specified page size.
 * @param {number} page - The page number to fetch.
 * @param {number} pageSize - The number of workouts per page.
 * @returns {Promise<Object>} The API response data.
 */
async function fetchWorkoutsPage(page, pageSize) {
  try {
    const response = await axios.get('https://api.hevyapp.com/v1/workouts', {
      params: {
        page: page,
        pageSize: pageSize
      },
      headers: {
        'accept': 'application/json',
        'api-key': API_KEY
      }
    });
    return response.data;
  } catch (error) {
    console.error(`Error fetching workouts on page ${page}:`, error.message);
    process.exit(1);
  }
}

/**
 * Formats a workout into multiple rows—one per exercise set.
 * Each row includes:
 *   - Workout ID
 *   - Workout Title
 *   - Workout Date (using created_at)
 *   - Exercise Title
 *   - Set Index
 *   - Weight (kg)
 *   - Reps
 *   - Duration (seconds)
 *   - Custom Metric (if any)
 *
 * @param {Object} workout - A workout object from the API.
 * @returns {Array} Array of formatted rows.
 */
function formatWorkoutSets(workout) {
  let rows = [];
  const workoutId = workout.id || "N/A";
  const workoutTitle = workout.title || "N/A";
  const workoutDate = workout.created_at || "N/A";
  
  if (Array.isArray(workout.exercises)) {
    workout.exercises.forEach(exercise => {
      const exerciseTitle = exercise.title || "N/A";
      if (Array.isArray(exercise.sets)) {
        exercise.sets.forEach((set, index) => {
          rows.push([
            workoutId,
            workoutTitle,
            workoutDate,
            exerciseTitle,
            index,
            set.weight_kg || "N/A",
            set.reps || "N/A",
            set.duration_seconds || "N/A",
            set.custom_metric || "N/A"
          ]);
        });
      }
    });
  }
  return rows;
}

/**
 * Main function to fetch enough pages to accumulate at least 20 workouts,
 * then format and display their exercise set details.
 */
async function main() {
  const pageSize = 9;
  let page = 1;
  let workouts = [];
  let totalPages = 1;

  // Continue fetching pages until we have at least 20 workouts or run out of pages.
  while (workouts.length < 20 && page <= totalPages) {
    const data = await fetchWorkoutsPage(page, pageSize);
    // Expected structure: { page: 1, page_count: X, workouts: [ ... ] }
    totalPages = data.page_count || 1;
    workouts = workouts.concat(data.workouts || []);
    page++;
  }

  // Slice to the first 20 workouts (if there are more)
  const topWorkouts = workouts.slice(0, 20);

  // Format each workout's sets into rows.
  let allRows = [];
  topWorkouts.forEach(workout => {
    const rows = formatWorkoutSets(workout);
    allRows = allRows.concat(rows);
  });
  
  console.log("Formatted Completed Workout Sets:");
  console.table(allRows);
}

main();