package model

data class BusRelation(
    private var id: Int = 0,
    private val companyName: String,
    private val startPoint: String,
    private val endPoint: String,
    private val departureTime: String,
    private val arrivalTime: String,
    private val note: String
) {
    fun toJson(): String {
        return "{ \"id\": ${id},\n" +
                "\"companyName\": \"$companyName\",\n" +
                "\"startPoint\": \"$startPoint\",\n" +
                "\"endPoint\": \"$endPoint\",\n" +
                "\"departureTime\": \"$departureTime\",\n" +
                "\"arrivalTime\": \"$arrivalTime\",\n" +
                "\"note\": \"$note\" },\n"
    }
}