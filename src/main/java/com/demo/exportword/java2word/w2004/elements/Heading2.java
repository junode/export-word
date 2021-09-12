package com.demo.exportword.java2word.w2004.elements;


import com.demo.exportword.java2word.interfaces.IFluentElement;
import com.demo.exportword.java2word.w2004.style.HeadingStyle;

public class Heading2 extends AbstractHeading<HeadingStyle> implements IFluentElement<Heading2> {

    //Constructor
    private Heading2(String value){
        super("Heading2", value);
    }

    /***
     * @param The value of the paragraph
     * @return the Fluent @Heading2
     */
    public static Heading2 with(String string) {
        return new Heading2(string);
    }


    @Override
    public Heading2 create() {
        return this;
    }

    // #### Getters and setters ####

}
