package org.example.model;

public record CrewAvailabilitySummary(
        int openerCount,
        int closerCount,
        int maximumAvailableCrewCount,
        Integer maximumAvailableHour,
        int minimumAvailableCrewCount,
        Integer minimumAvailableHour
) {
}
