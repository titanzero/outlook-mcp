/**
 * Meeting rooms / Places tools for Outlook MCP server
 */
const { getGraphClient, graphGetPaginated } = require('../utils/graph-client');
const { isAuthError, makeErrorResponse, makeResponse } = require('../utils/response-helpers');
const { formatResponse } = require('../utils/response-formatter');
const { z } = require('zod');

/**
 * List meeting rooms (places) in the organization
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListRooms(args) {
  const { roomList, count = 100 } = args;
  
  try {
    const client = await getGraphClient();
    
    let rooms = [];
    
    if (roomList) {
      // Get rooms from a specific room list
      const response = await graphGetPaginated(client, `places/microsoft.graph.roomList/${roomList}/rooms`, {
        $select: 'id,displayName,emailAddress,capacity,building,label,floorNumber,phone,address',
        $top: Math.min(count, 100),
      });
      rooms = response.value || [];
    } else {
      // Get all room lists first, then get rooms from each
      const roomListsResponse = await client.api('places/microsoft.graph.roomList').get();
      const roomLists = roomListsResponse.value || [];
      
      for (const list of roomLists) {
        try {
          const listRoomsResponse = await graphGetPaginated(client, `places/microsoft.graph.roomList/${list.id}/rooms`, {
            $select: 'id,displayName,emailAddress,capacity,building,label,floorNumber,phone,address',
            $top: Math.min(count, 100),
          });
          rooms.push(...(listRoomsResponse.value || []));
        } catch (e) {
          // Skip lists we can't access
        }
      }
    }
    
    const structured = rooms.map(r => ({
      id: r.id,
      name: r.displayName,
      email: r.emailAddress,
      capacity: r.capacity,
      building: r.building,
      floor: r.floorNumber,
      label: r.label,
      phone: r.phone,
      address: r.address ? {
        street: r.address.street,
        city: r.address.city,
        state: r.address.state,
        postalCode: r.address.postalCode,
        country: r.address.countryOrRegion,
      } : null,
    }));
    
    const textFallback = [
      `Meeting Rooms (${rooms.length})${roomList ? ` from list ${roomList}` : ''}:`,
      '',
      rooms.length === 0
        ? '  No rooms found. Your organization may not have room lists configured.'
        : rooms.map((r, i) => {
            const addr = r.address ? `, ${r.address.street}, ${r.address.city}` : '';
            return `${i + 1}. ${r.name} (${r.email})${r.capacity ? ` - Capacity: ${r.capacity}` : ''}${r.building ? `, Building: ${r.building}` : ''}${r.floor ? `, Floor: ${r.floor}` : ''}${addr}`;
          }).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ rooms: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Places/Rooms API not available. Your organization may not have meeting rooms configured in Microsoft 365.');
    }
    return makeErrorResponse(`Error listing rooms: ${error.message}`);
  }
}

/**
 * List room lists (groups of rooms)
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleListRoomLists(_args) {
  try {
    const client = await getGraphClient();
    
    const response = await client.api('places/microsoft.graph.roomList').get();
    const roomLists = response.value || [];
    
    const structured = roomLists.map(r => ({
      id: r.id,
      name: r.displayName,
      email: r.emailAddress,
      address: r.address ? {
        street: r.address.street,
        city: r.address.city,
        state: r.address.state,
        postalCode: r.address.postalCode,
        country: r.address.countryOrRegion,
      } : null,
    }));
    
    const textFallback = [
      `Room Lists (${roomLists.length}):`,
      '',
      roomLists.length === 0
        ? '  No room lists found.'
        : roomLists.map((r, i) => {
            const addr = r.address ? ` (${r.address.street}, ${r.address.city})` : '';
            return `${i + 1}. ${r.name} (${r.email})${addr} [id: ${r.id}]`;
          }).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse({ roomLists: structured }, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    if (error.statusCode === 404) {
      return makeErrorResponse('Places/Room Lists API not available. Your organization may not have meeting rooms configured in Microsoft 365.');
    }
    return makeErrorResponse(`Error listing room lists: ${error.message}`);
  }
}

/**
 * Check room availability for a time range
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleCheckRoomAvailability(args) {
  const { roomEmail, startTime, endTime } = args;
  
  if (!roomEmail) {
    return makeErrorResponse('Room email address is required (from list-rooms).');
  }
  if (!startTime || !endTime) {
    return makeErrorResponse('startTime and endTime are required (ISO 8601 format).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Use getSchedule to check room availability
    const response = await client.api('me/calendar/getSchedule').post({
      schedules: [roomEmail],
      startTime: { dateTime: startTime, timeZone: 'UTC' },
      endTime: { dateTime: endTime, timeZone: 'UTC' },
      availabilityViewInterval: 30,
    });
    
    const schedule = response.value?.[0];
    if (!schedule) {
      return makeErrorResponse(`No schedule information found for ${roomEmail}.`);
    }
    
    const structured = {
      roomEmail,
      startTime,
      endTime,
      availabilityView: schedule.availabilityView,
      scheduleItems: schedule.scheduleItems?.map(item => ({
        start: item.start?.dateTime,
        end: item.end?.dateTime,
        status: item.status,
        subject: item.subject,
        location: item.location,
        isPrivate: item.isPrivate,
      })) || [],
      workingHours: schedule.workingHours,
    };
    
    // Parse availability view (0=free, 1=tentative, 2=busy, 3=oof, 4=working elsewhere)
    const statusMap = { '0': 'Free', '1': 'Tentative', '2': 'Busy', '3': 'Out of Office', '4': 'Working Elsewhere' };
    const view = schedule.availabilityView || '';
    const slots = view.length;
    const intervalMinutes = 30;
    
    const textFallback = [
      `Room Availability: ${roomEmail}`,
      `Time Range: ${startTime} to ${endTime}`,
      `Interval: ${intervalMinutes} minutes (${slots} slots)`,
      '',
      `Availability View: ${view}`,
      'Legend: 0=Free, 1=Tentative, 2=Busy, 3=OOF, 4=Working Elsewhere',
      '',
      'Busy Periods:',
      schedule.scheduleItems?.length > 0
        ? schedule.scheduleItems.map(item => `- ${item.start?.dateTime} to ${item.end?.dateTime}: ${item.subject || 'Busy'} (${item.status})`).join('\n')
        : '  No conflicts - room appears available',
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error checking room availability: ${error.message}`);
  }
}

/**
 * Find available rooms for a time range
 * @param {object} args - Tool arguments
 * @returns {object} - MCP response
 */
async function handleFindAvailableRooms(args) {
  const { startTime, endTime, capacity, roomList, count = 20 } = args;
  
  if (!startTime || !endTime) {
    return makeErrorResponse('startTime and endTime are required (ISO 8601 format).');
  }
  
  try {
    const client = await getGraphClient();
    
    // Get rooms to check
    let rooms = [];
    if (roomList) {
      const response = await graphGetPaginated(client, `places/microsoft.graph.roomList/${roomList}/rooms`, {
        $select: 'id,displayName,emailAddress,capacity,building,label',
        $top: 100,
      });
      rooms = response.value || [];
    } else {
      const roomListsResponse = await client.api('places/microsoft.graph.roomList').get();
      const roomLists = roomListsResponse.value || [];
      
      for (const list of roomLists) {
        try {
          const listRoomsResponse = await graphGetPaginated(client, `places/microsoft.graph.roomList/${list.id}/rooms`, {
            $select: 'id,displayName,emailAddress,capacity,building,label',
            $top: 100,
          });
          rooms.push(...(listRoomsResponse.value || []));
        } catch (e) {
          // Skip lists we can't access
        }
      }
    }
    
    // Filter by capacity if specified
    if (capacity) {
      rooms = rooms.filter(r => r.capacity && r.capacity >= capacity);
    }
    
    // Check availability for each room (batch in groups of 20 due to API limits)
    const availableRooms = [];
    const batchSize = 20;
    
    for (let i = 0; i < Math.min(rooms.length, count * 2); i += batchSize) {
      const batch = rooms.slice(i, i + batchSize);
      const schedules = batch.map(r => r.emailAddress);
      
      try {
        const response = await client.api('me/calendar/getSchedule').post({
          schedules,
          startTime: { dateTime: startTime, timeZone: 'UTC' },
          endTime: { dateTime: endTime, timeZone: 'UTC' },
          availabilityViewInterval: 60,
        });
        
        for (const schedule of response.value || []) {
          const room = batch.find(r => r.emailAddress === schedule.scheduleId);
          if (room) {
            // Check if room is free during the entire period
            const hasConflict = (schedule.scheduleItems || []).some(item => 
              item.status === 'busy' || item.status === 'oof'
            );
            
            if (!hasConflict) {
              availableRooms.push({
                id: room.id,
                name: room.displayName,
                email: room.emailAddress,
                capacity: room.capacity,
                building: room.building,
                label: room.label,
              });
            }
          }
        }
      } catch (e) {
        // Skip batch on error
      }
    }
    
    const structured = {
      startTime,
      endTime,
      capacityFilter: capacity || null,
      availableRooms: availableRooms.slice(0, count),
      totalChecked: rooms.length,
    };
    
    const textFallback = [
      `Available Rooms (${availableRooms.length} of ${rooms.length} checked)`,
      `Time: ${startTime} to ${endTime}${capacity ? `, Min Capacity: ${capacity}` : ''}`,
      '',
      availableRooms.length === 0
        ? '  No available rooms found for this time range.'
        : availableRooms.slice(0, count).map((r, i) => 
            `${i + 1}. ${r.name} (${r.email})${r.capacity ? ` - Cap: ${r.capacity}` : ''}${r.building ? `, Bldg: ${r.building}` : ''}${r.label ? `, ${r.label}` : ''}`
          ).join('\n'),
    ].join('\n');
    
    return makeResponse(formatResponse(structured, textFallback));
  } catch (error) {
    if (isAuthError(error)) {
      return makeErrorResponse(error.message);
    }
    return makeErrorResponse(`Error finding available rooms: ${error.message}`);
  }
}

module.exports = {
  handleListRooms,
  handleListRoomLists,
  handleCheckRoomAvailability,
  handleFindAvailableRooms,
};