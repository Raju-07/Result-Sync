### Purpose

This file purpose is to log all the key changes i'm doing to improve it performance and optimazing. Also trying to add more features.

### Changes

#### self.process_data:

***before***

this was a list earlier `self.process_data: list` and we're performing operation such as `lookup` and `append` . reason to update is operation `lookup` was taking time of `O(n`) for each student.

***after***

we changed the data type of `process_data` from `list` to `set` which is much faster in lookup i.e. in time `O(1)` (constant time) and replace the `append` operation with `add`.

***result***

Time Complexity :  `O(n)`  --> `O(1)`  # for lookup operation in set

Space Complexity: `O(n)` # same as before

#### data:

changes the data type `data` from `list` to `tuple` as semantic improvement becuase all the data inside `data` eliments are fixed.

#### col_header:

changes the data type `col_header` from `list` to `tuple` as semantic improvement becuase all the data inside `col_header` eliments are fixed.

#### Calculate time Function:

created a function to calcuate the estimate time for showing better estimate time.

```Python
from datetime import datetime

#starting time
starting_time = int(datetime.now().strftime('%S'))
# process time consume
end_time = int(datetime.now().strftime('%S'))

def calculate_time(start_time: int = self.starting_time, end_time: int = self.end_time) -> int:
    if end_time >= start_time:
        # normal case: no wrap-around
        return (end_time - start_time) * self.total_record
    
    # edge case: wrap-around when end_time < start_time
    return ((60 - start_time) + end_time) * self.total_record

```
